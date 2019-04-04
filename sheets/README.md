# Sheets

Import yesterday's Mediahawk calls into Google Sheets

## Setup

1. Create a new Google Script
2. Add [Code.gs](Code.gs)
3. Set the config at the top of [Code.gs](Code.gs)
    * **USER** - Your Mediahawk API username, contact Mediahawk if you do not know this
    * **PASSWORD** - Your Mediahawk API password, contact Mediahawk if you do not know this
    * **SPREADSHEET_URL** - The full URL of the spreadsheet you want data inserted into, starting with “https” and ending with “/edit”
    * **SHEET_NAME** - The sheet you want data inseted to within the Google Sheet, this will be created if it does not exist
    * **API_FIELD** - An array of the fields you want inserted into your sheet in order, you can find the available fields in the *GetCallsByDate* method in the [API Documentation](https://support.mediahawk.co.uk/support/solutions/articles/17000053028-mediahawk-api-documentation)
4. Run the *main* method and follow the permission review
5. Set up a trigger to run the *main* method on a schedule

**Important** – This imports the calls from yesterday so it should run on a “Day timer”, it is best to set it to run “6am to 7am” to ensure all calls are available by the time they are added to your spreadsheet.