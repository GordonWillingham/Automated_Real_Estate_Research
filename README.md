# Charlotte, NC listing report (placeholder)

Because this environment does not have outbound internet access, I could not query real-estate portals or MLS feeds to gather current listings for north Charlotte inside I-485.

To keep the requested workbook deliverable reproducible, I generated `charlotte_listings.xlsx` with the expected columns and a placeholder row explaining the limitation. The `generate_listings_placeholder.py` script writes a minimal XLSX package without external dependencies; rerun it to regenerate the workbook.

If you rerun this in an environment with internet access, you can replace `ROWS` in `generate_listings_placeholder.py` with live listing data (sorted by a computed Resale Value Score) and re-execute the script to refresh the file.
