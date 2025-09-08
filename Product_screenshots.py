import nest_asyncio
import asyncio
import pandas as pd
import datetime
import calendar
import os
from playwright.async_api import async_playwright

nest_asyncio.apply()

# Path to Excel file
excel_path = r"Path to Excel file where code can read the link "

# Base folder for screenshots
base_folder = r"Base folder for screenshots"

# Read Excel (assumes columns: Brand, Model, FK link, RD link, Croma link)
df = pd.read_excel(excel_path)

# Current date info
today = datetime.date.today()
month_name = calendar.month_name[today.month]
week_num = today.isocalendar()[1]
day_str = today.strftime("%d-%b-%Y")

async def capture_screenshot(site_name, url, page, save_path):
    try:
        if not url or not isinstance(url, str) or not url.startswith("http"):
            print(f"❌ Skipping invalid URL: {url}")
            return
        await page.goto(url, timeout=60000, wait_until="domcontentloaded")
        await page.wait_for_timeout(5000)  # wait extra for dynamic content
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        await page.screenshot(path=save_path, full_page=True)
        print(f"✅ Saved: {save_path}")
    except Exception as e:
        print(f"❌ Error with {site_name} URL {url}: {e}")

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
            locale="en-US",
            java_script_enabled=True
        )
        page = await context.new_page()

        for _, row in df.iterrows():
            brand = str(row.get("Brand", "Unknown")).strip()
            model = str(row.get("Model", "Unknown")).strip()

            for site_col, site_name in [("FK link", "FK"), ("RD link", "RD"), ("Croma link", "Croma")]:
                url = row.get(site_col, None)
                if not url or not isinstance(url, str) or not url.startswith("http"):
                    continue

                # Build folder structure
                folder_path = os.path.join(
                    base_folder,
                    month_name,
                    f"Week-{week_num}",
                    day_str,
                    brand,
                    model,
                    site_name
                )
                filename = f"{brand}_{model}_{site_name}.png".replace(" ", "_")
                save_path = os.path.join(folder_path, filename)

                await capture_screenshot(site_name, url, page, save_path)

        await browser.close()

# Run the script
await main()
