# Warehouse Pump Truck & Asset Register

Mobile-first Google Apps Script web app for:
- Individually registering pump trucks with automatic IDs (`PT-0001`, `PT-0002`, ...)
- Recording pump truck type, condition, area, tags, notes and a phone photo
- Dashboard totals by condition
- General warehouse asset counting by asset type, quantity, area, tags, notes and photo
- History, search, filters, edit/delete for pump trucks, and delete for asset count records
- Photos saved into the supplied Google Drive folder, inside `Pump Trucks` and `Asset Counter` subfolders
- Data saved into the supplied Google Spreadsheet
- Automatic creation/formatting of the blank spreadsheet tabs

## Google Sheet
Spreadsheet ID:
`1kpkk1UdmMExtew7s504aPZTDwKQvWjhxt83sdaijis`

Created/used sheets:
- `Pump Trucks`
- `Asset Counts`
- `Asset Types`
- `Settings`

## Google Drive
Folder ID:
`1rqHbMgyFSKYRDbxv-UHFpby5S-9-Wlz`

The script creates two subfolders automatically:
- `Pump Trucks`
- `Asset Counter`

## Setup
1. Create a new Google Apps Script project at script.google.com.
2. Add these files with the exact names: `Code.gs`, `Index.html`, `Styles.html`, `JavaScript.html`.
3. Paste the contents from this package into each corresponding file.
4. Save the project.
5. Run `setupProject()` once from the Apps Script editor. Google will ask for permission to use Sheets/Drive; approve it.
6. Deploy > New deployment > Web app.
7. Set **Execute as:** Me.
8. Set **Who has access:** choose the option appropriate for your warehouse (for a private internal tool, restrict it to your Google account/domain where possible).
9. Open the deployment URL on the phone.

## First run
The `setupProject()` function creates the sheet headers and default asset types if the tabs are empty. The app also calls setup automatically when it first loads.

## Notes
- Photo capture uses `<input type=file accept=image/* capture=environment>`, so supported mobile browsers can open the phone camera.
- Photos are resized in the browser before upload to reduce file size.
- The app uses Apps Script's `google.script.run`, so it should be used from the Apps Script web-app deployment rather than directly opened as a local HTML file.
- The image URL stored in the Sheet is a Drive thumbnail URL. Users still need permission to the underlying Drive files/folder.
