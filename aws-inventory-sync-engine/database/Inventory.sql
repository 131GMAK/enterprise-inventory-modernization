-- ============================================================
-- INVENTORY DATABASE — FULL SCHEMA
-- Consolidated from all CREATE TABLE, INSERT, index, constraint,
-- and trigger statements. Safe to run top-to-bottom on a fresh
-- PostgreSQL database to fully recreate this schema.
-- ============================================================

-- ============================================================
-- 1. CATEGORIES
-- ============================================================
CREATE TABLE categories (
    category_id     BIGSERIAL PRIMARY KEY,
    parent_id       BIGINT REFERENCES categories(category_id),
    name            TEXT NOT NULL,
    slug            TEXT UNIQUE NOT NULL,
    path            TEXT,
    created_at      TIMESTAMPTZ DEFAULT now()
);

-- 1. Insert Top-Level Parent Categories
INSERT INTO categories (name, slug) VALUES
    ('Children’s Toys', 'childrens-toys'),
    ('Children Educational Items', 'children-educational-items'),
    ('Party Souvenirs', 'party-souvenirs'),
    ('Back to School', 'back-to-school');

-- 2. Insert Subcategories linked to Parent Category IDs
INSERT INTO categories (name, slug, parent_id) VALUES
    ('Wooden Toys', 'wooden-toys', (SELECT category_id FROM categories WHERE slug = 'childrens-toys')),
    ('Costumes', 'costumes', (SELECT category_id FROM categories WHERE slug = 'childrens-toys')),
    ('Montessori Educational Items', 'montessori-educational-items', (SELECT category_id FROM categories WHERE slug = 'children-educational-items')),
    ('Party Favors', 'party-favors', (SELECT category_id FROM categories WHERE slug = 'party-souvenirs')),
    ('Balloon', 'balloon', (SELECT category_id FROM categories WHERE slug = 'party-souvenirs')),
    ('Leisure Books', 'leisure-books', (SELECT category_id FROM categories WHERE slug = 'back-to-school'));

-- ============================================================
-- 2. PRODUCTS (Parent product)
-- ============================================================
CREATE TABLE products (
    product_id      BIGSERIAL PRIMARY KEY,
    category_id     BIGINT NOT NULL REFERENCES categories(category_id),
    name            TEXT NOT NULL,
    description     TEXT,
    brand           TEXT,
    base_sku        TEXT UNIQUE,
    is_active       BOOLEAN DEFAULT true,
    created_at      TIMESTAMPTZ DEFAULT now(),
    updated_at      TIMESTAMPTZ DEFAULT now()
);

-- ============================================================
-- 3. PRODUCT VARIANTS (Colour, Size, etc. - the actual sellable item)
-- ============================================================
CREATE TABLE product_variants (
    variant_id      BIGSERIAL PRIMARY KEY,
    product_id      BIGINT NOT NULL REFERENCES products(product_id),
    sku             TEXT UNIQUE NOT NULL,
    barcode         TEXT UNIQUE,
    attributes      JSONB NOT NULL DEFAULT '{}',   -- e.g. {"color": "Red", "size": "Large"}
    price_retail    NUMERIC(12,2),
    price_wholesale NUMERIC(12,2),
    weight_kg       NUMERIC(8,3),
    is_active       BOOLEAN DEFAULT true,
    created_at      TIMESTAMPTZ DEFAULT now(),
    updated_at      TIMESTAMPTZ DEFAULT now()
);

-- ============================================================
-- 4. PRODUCT IMAGES
-- ============================================================
CREATE TABLE product_images (
    image_id        BIGSERIAL PRIMARY KEY,
    product_id      BIGINT REFERENCES products(product_id),
    variant_id      BIGINT REFERENCES product_variants(variant_id),
    url             TEXT NOT NULL,
    is_primary      BOOLEAN DEFAULT false,
    sort_order      INT DEFAULT 0
);

-- ============================================================
-- 5. LOCATIONS (Warehouse + Main Shop)
-- ============================================================
CREATE TABLE locations (
    location_id     SMALLSERIAL PRIMARY KEY,
    code            TEXT UNIQUE NOT NULL,          -- 'WH', 'SHOP'
    name            TEXT NOT NULL,
    is_active       BOOLEAN DEFAULT true
);

-- Insert the two locations
INSERT INTO locations (code, name) VALUES
    ('WH', 'Warehouse'),
    ('SHOP', 'Main Shop');

-- ============================================================
-- 6. CURRENT STOCK LEVELS
-- ============================================================
CREATE TABLE inventory_levels (
    location_id         SMALLINT NOT NULL REFERENCES locations(location_id),
    variant_id          BIGINT NOT NULL REFERENCES product_variants(variant_id),
    quantity_on_hand    INTEGER NOT NULL DEFAULT 0,
    quantity_reserved   INTEGER NOT NULL DEFAULT 0,
    quantity_available  INTEGER GENERATED ALWAYS AS (quantity_on_hand - quantity_reserved) STORED,
    last_movement_at    TIMESTAMPTZ,
    updated_at          TIMESTAMPTZ DEFAULT now(),
    PRIMARY KEY (location_id, variant_id)
);

-- ============================================================
-- 7. STOCK MOVEMENTS (Source of truth)
-- ============================================================
CREATE TABLE stock_movements (
    movement_id     BIGSERIAL,
    location_id     SMALLINT NOT NULL,
    variant_id      BIGINT NOT NULL,
    movement_type   TEXT NOT NULL,                 -- RECEIVE, SALE, TRANSFER_OUT, TRANSFER_IN, ADJUSTMENT, RETURN, RESERVE, UNRESERVE
    quantity        INTEGER NOT NULL,
    reference_type  TEXT,
    reference_id    TEXT,
    notes           TEXT,
    created_by      TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    client_id       UUID,
    PRIMARY KEY (movement_id, created_at)
) PARTITION BY RANGE (created_at);

CREATE TABLE stock_movements_default 
PARTITION OF stock_movements DEFAULT;

-- Partitions (monthly ranges)
CREATE TABLE stock_movements_2026_09 PARTITION OF stock_movements
    FOR VALUES FROM ('2026-09-01') TO ('2026-10-01');

CREATE TABLE stock_movements_2026_10 PARTITION OF stock_movements
    FOR VALUES FROM ('2026-10-01') TO ('2026-11-01');

CREATE TABLE stock_movements_2026_11 PARTITION OF stock_movements
    FOR VALUES FROM ('2026-11-01') TO ('2026-12-01');

CREATE TABLE stock_movements_2026_12 PARTITION OF stock_movements
    FOR VALUES FROM ('2026-12-01') TO ('2027-01-01');

CREATE TABLE stock_movements_2027_01 PARTITION OF stock_movements
    FOR VALUES FROM ('2027-01-01') TO ('2027-02-01');

-- NOTE: there is no partition covering August 2026 (today's date).
-- Any stock movement inserted before September 1, 2026 will fail
-- with "no partition of relation found for row" unless one is added
-- (e.g. stock_movements_2026_08) or a DEFAULT partition is created
-- as a catch-all safety net.

-- ============================================================
-- 8. STOCK TRANSFERS (Warehouse <-> Shop)
-- ============================================================
CREATE TABLE stock_transfers (
    transfer_id     BIGSERIAL PRIMARY KEY,
    from_location   SMALLINT NOT NULL REFERENCES locations(location_id),
    to_location     SMALLINT NOT NULL REFERENCES locations(location_id),
    status          TEXT NOT NULL DEFAULT 'PENDING',
    created_at      TIMESTAMPTZ DEFAULT now(),
    completed_at    TIMESTAMPTZ
);

-- ============================================================
-- STOCK TRANSFER LINES (line items per transfer)
-- ============================================================
CREATE TABLE stock_transfer_lines (
    transfer_id     BIGINT REFERENCES stock_transfers(transfer_id),
    variant_id      BIGINT REFERENCES product_variants(variant_id),
    quantity        INTEGER NOT NULL,
    PRIMARY KEY (transfer_id, variant_id)
);

-- ============================================================
-- 9. OFFLINE SYNC QUEUE
-- ============================================================
CREATE TABLE sync_queue (
    queue_id        BIGSERIAL PRIMARY KEY,
    device_id       TEXT NOT NULL,
    payload         JSONB NOT NULL,
    client_id       UUID NOT NULL UNIQUE,
    created_at      TIMESTAMPTZ DEFAULT now(),
    processed_at    TIMESTAMPTZ,
    status          TEXT NOT NULL DEFAULT 'PENDING'
                        CHECK (status IN ('PENDING', 'PROCESSED', 'FAILED')),
    error_message   TEXT
);

-- ============================================================
-- 10. SHIPMENT BATCHES (for bulk imports)
-- ============================================================
CREATE TABLE shipment_batches (
    shipment_id         BIGSERIAL PRIMARY KEY,
    batch_code          TEXT UNIQUE NOT NULL,         -- e.g., 'TINA_ORDER_2026_06'
    naira_per_usd       NUMERIC(10,2) NOT NULL,       -- 1470
    rmb_per_usd         NUMERIC(10,2) NOT NULL,       -- 7
    total_freight_clr   NUMERIC(14,2) NOT NULL,       -- 28500000
    imported_at         TIMESTAMPTZ DEFAULT now()
);

-- ============================================================
-- 11. SHIPMENT ITEMS (for bulk imports)
-- ============================================================
CREATE TABLE shipment_items (
    shipment_item_id    BIGSERIAL PRIMARY KEY,
    shipment_id         BIGINT NOT NULL REFERENCES shipment_batches(shipment_id),
    variant_id          BIGINT NOT NULL REFERENCES product_variants(variant_id),
    ctns                INTEGER NOT NULL,
    qty                 INTEGER NOT NULL,
    unit_price_rmb      NUMERIC(12,2) NOT NULL,
    total_price_rmb     NUMERIC(12,2) NOT NULL,
    cbm                 NUMERIC(8,4),
    length_cm           NUMERIC(6,2),
    width_cm            NUMERIC(6,2),
    height_cm           NUMERIC(6,2),
    naira_cost_china    NUMERIC(12,2),
    freight_per_unit    NUMERIC(12,2),
    landed_wh_price     NUMERIC(12,2),
    lcl_selling_price   NUMERIC(12,2),
    gross_sale_value    NUMERIC(14,2),
    profitability       NUMERIC(12,2)
);

-- ============================================================
-- INDEXES
-- ============================================================

-- Categories
CREATE INDEX idx_categories_parent ON categories(parent_id);
CREATE INDEX idx_categories_path ON categories (path text_pattern_ops);

-- Products
CREATE INDEX idx_products_category ON products(category_id);

-- Product variants
CREATE INDEX idx_variants_product ON product_variants(product_id);
CREATE INDEX idx_variants_attributes ON product_variants USING GIN (attributes);
-- Note: no separate indexes needed on sku/barcode — their UNIQUE
-- constraints already create equivalent indexes automatically.

-- Product images
CREATE INDEX idx_images_product ON product_images(product_id);
CREATE INDEX idx_images_variant ON product_images(variant_id);

-- Inventory levels
CREATE INDEX idx_inventory_levels_variant ON inventory_levels(variant_id);
CREATE INDEX idx_inventory_levels_last_movement ON inventory_levels(last_movement_at);

-- Stock movements
CREATE INDEX idx_movements_variant_location ON stock_movements(variant_id, location_id);
CREATE INDEX idx_movements_created_at ON stock_movements(created_at);
CREATE INDEX idx_movements_reference ON stock_movements(reference_type, reference_id);
CREATE UNIQUE INDEX idx_movements_client_id ON stock_movements(client_id, created_at);

-- Sync queue
CREATE INDEX idx_sync_queue_pending ON sync_queue (created_at) WHERE status = 'PENDING';

-- Shipment batches/items
CREATE INDEX idx_shipment_items_shipment ON shipment_items(shipment_id);
CREATE INDEX idx_shipment_items_variant ON shipment_items(variant_id);
CREATE INDEX idx_shipment_batches_imported ON shipment_batches(imported_at);

-- ============================================================
-- CONSTRAINTS (added after initial table creation)
-- ============================================================

-- Categories
ALTER TABLE categories
    ADD CONSTRAINT chk_not_self_parent CHECK (parent_id <> category_id);

-- Product variants
ALTER TABLE product_variants
    ADD CONSTRAINT chk_price_retail_nonneg CHECK (price_retail >= 0),
    ADD CONSTRAINT chk_price_wholesale_nonneg CHECK (price_wholesale >= 0),
    ADD CONSTRAINT chk_weight_nonneg CHECK (weight_kg >= 0);

-- Product images
ALTER TABLE product_images
    ADD CONSTRAINT chk_product_or_variant CHECK (product_id IS NOT NULL OR variant_id IS NOT NULL);

-- Inventory levels
ALTER TABLE inventory_levels
    ADD CONSTRAINT chk_qty_on_hand_nonneg CHECK (quantity_on_hand >= 0),
    ADD CONSTRAINT chk_qty_reserved_nonneg CHECK (quantity_reserved >= 0),
    ADD CONSTRAINT chk_reserved_lte_on_hand CHECK (quantity_reserved <= quantity_on_hand);

-- Stock movements
ALTER TABLE stock_movements
    ADD CONSTRAINT fk_movements_location FOREIGN KEY (location_id) REFERENCES locations(location_id),
    ADD CONSTRAINT fk_movements_variant FOREIGN KEY (variant_id) REFERENCES product_variants(variant_id),
    ADD CONSTRAINT chk_movement_type CHECK (movement_type IN
        ('RECEIVE','SALE','TRANSFER_OUT','TRANSFER_IN','ADJUSTMENT','RETURN','RESERVE','UNRESERVE'));

-- Stock transfers
ALTER TABLE stock_transfers
    ADD CONSTRAINT chk_diff_locations CHECK (from_location <> to_location),
    ADD CONSTRAINT chk_transfer_status CHECK (status IN ('PENDING','IN_TRANSIT','COMPLETED','CANCELLED'));

-- Stock transfer lines
ALTER TABLE stock_transfer_lines
    ADD CONSTRAINT chk_transfer_qty_positive CHECK (quantity > 0);


-- Shipment batches
ALTER TABLE shipment_batches
    ADD CONSTRAINT chk_naira_rate_pos CHECK (naira_per_usd > 0),
    ADD CONSTRAINT chk_rmb_rate_pos CHECK (rmb_per_usd > 0),
    ADD CONSTRAINT chk_freight_clr_nonneg CHECK (total_freight_clr >= 0);

-- Shipment items
ALTER TABLE shipment_items
    ADD CONSTRAINT chk_shipment_items_qty_pos CHECK (qty > 0),
    ADD CONSTRAINT chk_shipment_items_ctns_pos CHECK (ctns > 0),
    ADD CONSTRAINT chk_unit_price_rmb_nonneg CHECK (unit_price_rmb >= 0),
    ADD CONSTRAINT chk_landed_wh_nonneg CHECK (landed_wh_price >= 0);

-- ============================================================
-- TRIGGER FUNCTIONS
-- ============================================================

-- Generic "touch updated_at on row update" function,
-- reused across products and product_variants.
CREATE OR REPLACE FUNCTION set_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_products_updated_at
BEFORE UPDATE ON products
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

CREATE TRIGGER trg_variants_updated_at
BEFORE UPDATE ON product_variants
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

-- Note: inventory_levels is NOT auto-updated by a stock_movements
-- trigger in this schema — that sync is handled at the application
-- layer (Odoo offline POS tracker), by design.

-- ============================================================
-- END OF SCHEMA
-- ============================================================