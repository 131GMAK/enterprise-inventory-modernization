import os
import pandas as pd
import psycopg2
from dotenv import load_dotenv

load_dotenv()

# Database Connection
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "inventory_db")
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASSWORD = os.getenv("DB_PASSWORD", "Regular$hosql25")

def get_db_connection():
    return psycopg2.connect(
        host=DB_HOST,
        port=DB_PORT,
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD
    )

def parse_and_ingest():
    # Dynamic absolute path targeting data/sample.xlsx
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.normpath(os.path.join(script_dir, "..", "data", "sample.xlsx"))

    print(f"Reading file: {file_path}")
    raw_df = pd.read_excel(file_path, header=None)

    # 1. Extract Batch Metadata from Row 0
    batch_code = str(raw_df.iloc[0, 1]).strip().replace(" ", "_")
    naira_per_usd = float(raw_df.iloc[0, 22]) if pd.notna(raw_df.iloc[0, 22]) else 1470.0
    rmb_per_usd = float(raw_df.iloc[0, 23]) if pd.notna(raw_df.iloc[0, 23]) else 7.0
    total_freight_clr = float(raw_df.iloc[0, 25]) if pd.notna(raw_df.iloc[0, 25]) else 0.0

    print(f"Batch Code: {batch_code} | NGN/$: {naira_per_usd} | Freight/Clr: NGN {total_freight_clr:,.2f}")

    # 2. Extract Data Rows (Row 2 onwards)
    data_df = raw_df.iloc[2:].copy()
    
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            # Step A: Upsert Parent Shipment Batch
            cur.execute("""
                INSERT INTO shipment_batches (batch_code, naira_per_usd, rmb_per_usd, total_freight_clr)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (batch_code) DO UPDATE 
                SET naira_per_usd = EXCLUDED.naira_per_usd,
                    rmb_per_usd = EXCLUDED.rmb_per_usd,
                    total_freight_clr = EXCLUDED.total_freight_clr
                RETURNING shipment_id;
            """, (batch_code, naira_per_usd, rmb_per_usd, total_freight_clr))
            shipment_id = cur.fetchone()[0]

            # Step B: Retrieve Default Category ('childrens-toys')
            cur.execute("SELECT category_id FROM categories WHERE slug = 'childrens-toys';")
            cat_row = cur.fetchone()
            if cat_row:
                category_id = cat_row[0]
            else:
                # Fallback to first available category
                cur.execute("SELECT category_id FROM categories ORDER BY category_id ASC LIMIT 1;")
                category_id = cur.fetchone()[0]

            # Step C: Retrieve or Create Default Warehouse Location ('WH')
            cur.execute("SELECT location_id FROM locations WHERE code = 'WH';")
            loc_row = cur.fetchone()
            if loc_row:
                location_id = loc_row[0]
            else:
                cur.execute("""
                    INSERT INTO locations (name, code, type) 
                    VALUES ('Main Warehouse', 'WH', 'WAREHOUSE') 
                    RETURNING location_id;
                """)
                location_id = cur.fetchone()[0]

            # Step D: Iterate and Ingest Line Items
            items_processed = 0
            for idx, row in data_df.iterrows():
                # 1. Combine all non-empty row text to check for summary keywords
                row_text = " ".join([str(val) for val in row.values if pd.notna(val)]).upper()
                
                # 2. Extract item number candidate from column index 3
                raw_item_no = str(row[3]).strip() if pd.notna(row[3]) else ""
                
                # 3. GUARD CLAUSE: Skip empty rows and container summary/footer rows
                summary_indicators = ["CBM FREIGHT", "TOTAL FREIGHT", "CONTAINER", "CTNS"]
                
                is_summary_row = any(indicator in row_text for indicator in summary_indicators)
                is_invalid_sku = len(raw_item_no) < 2 or "CTNS" in raw_item_no.upper()
                
                if not raw_item_no or is_summary_row or is_invalid_sku:
                    print(f"--> Skipping non-item summary/footer row at index {idx}: {row_text[:45]}...")
                    continue

                # --- RESTORED: Explicit raw_sku definition for downstream database insertions ---
                raw_sku = str(raw_item_no).strip()
                unique_sku = f"{batch_code}_{idx}_{raw_sku}"

                ctns = int(row[6]) if pd.notna(row[6]) else 1
                qty = int(row[7]) if pd.notna(row[7]) else 0
                price_rmb = float(row[8]) if pd.notna(row[8]) else 0.0
                total_price_rmb = float(row[9]) if pd.notna(row[9]) else (qty * price_rmb)
                cbm = float(row[11]) if pd.notna(row[11]) else 0.0

                # Master Carton Dimensions (Cols 13, 14, 15)
                length_cm = float(row[13]) if pd.notna(row[13]) else 0.0
                width_cm = float(row[14]) if pd.notna(row[14]) else 0.0
                height_cm = float(row[15]) if pd.notna(row[15]) else 0.0

                # Financial Economics
                lcl_selling_price = float(row[17]) if pd.notna(row[17]) else 0.0
                landed_wh_price = float(row[18]) if pd.notna(row[18]) else 0.0
                gross_sale_value = float(row[19]) if pd.notna(row[19]) else 0.0
                naira_cost_china = float(row[21]) if pd.notna(row[21]) else 0.0
                freight_unit = float(row[24]) if pd.notna(row[24]) else 0.0
                profitability = float(row[26]) if pd.notna(row[26]) else 0.0

                # 1. Product Master
                cur.execute("""
                    INSERT INTO products (category_id, name, base_sku)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (base_sku) DO UPDATE SET updated_at = now()
                    RETURNING product_id;
                """, (category_id, f"Import Item {raw_sku}", unique_sku))
                product_id = cur.fetchone()[0]

                # 2. Product Variant
                carton_attrs = f'{{"cbm": {cbm}, "master_carton": {{"length_cm": {length_cm}, "width_cm": {width_cm}, "height_cm": {height_cm}}}}}'
                cur.execute("""
                    INSERT INTO product_variants (product_id, sku, price_wholesale, price_retail, attributes)
                    VALUES (%s, %s, %s, %s, %s::jsonb)
                    ON CONFLICT (sku) DO UPDATE SET price_wholesale = EXCLUDED.price_wholesale
                    RETURNING variant_id;
                """, (product_id, unique_sku, landed_wh_price, lcl_selling_price, carton_attrs))
                variant_id = cur.fetchone()[0]

                # 3. Shipment Item
                cur.execute("""
                    INSERT INTO shipment_items (
                        shipment_id, variant_id, ctns, qty, unit_price_rmb, total_price_rmb,
                        cbm, length_cm, width_cm, height_cm, naira_cost_china,
                        freight_per_unit, landed_wh_price, lcl_selling_price,
                        gross_sale_value, profitability
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
                """, (
                    shipment_id, variant_id, ctns, qty, price_rmb, total_price_rmb,
                    cbm, length_cm, width_cm, height_cm, naira_cost_china,
                    freight_unit, landed_wh_price, lcl_selling_price,
                    gross_sale_value, profitability
                ))

                # 4. Warehouse Inventory Levels
                cur.execute("""
                    INSERT INTO inventory_levels (location_id, variant_id, quantity_on_hand)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (location_id, variant_id) DO UPDATE
                    SET quantity_on_hand = inventory_levels.quantity_on_hand + EXCLUDED.quantity_on_hand;
                """, (location_id, variant_id, qty))

                # 5. Stock Movements
                cur.execute("""
                    INSERT INTO stock_movements (location_id, variant_id, movement_type, quantity, reference_type, reference_id, notes)
                    VALUES (%s, %s, 'RECEIVE', %s, 'SHIPMENT_BATCH', %s, %s);
                """, (location_id, variant_id, qty, str(shipment_id), f"Import {batch_code}"))

                items_processed += 1

            conn.commit()
            print(f"Successfully processed {items_processed} items into inventory_db!")

    except Exception as e:
        conn.rollback()
        print(f"Error during ingestion: {e}")
        raise e
    finally:
        conn.close()

if __name__ == "__main__":
    parse_and_ingest()