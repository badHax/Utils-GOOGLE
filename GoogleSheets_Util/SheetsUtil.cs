using ConsoleTableExt;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.AppendRequest;

namespace GoogleSheets_Util
{

    public static class SheetsUtil
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = {
            SheetsService.Scope.SpreadsheetsReadonly,
            SheetsService.Scope.Drive, // if the file is in google drive
            SheetsService.Scope.DriveFile
        };
        static string ApplicationName = "Sheets API Util";

        /// <summary>
        /// authenticates and return an instance of the sheets service
        /// </summary>
        /// <returns></returns>
        public static SheetsService GetService()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static IList<object> GetCell(string sheetId, int row, string col)
        {
            // select the entire row
            string range = $"{col}{row}";
            // create the request with the sheet ID
            var request = GetService().Spreadsheets.Values.Get(sheetId, range);

            // get the data
            ValueRange response = request.Execute();

            // print the data
            if (response != null && response.Values.Count > 0)
            {
                // should just be one sheet
                return response.Values[0];
            }

            // return
            return null;
        }

        public static Sheet GetSheet(string spreadsheetId)
        {
            // create the request with the sheet ID
            var request = GetService().Spreadsheets.Get(spreadsheetId);
            // also include the sheet data
            request.IncludeGridData = true;

            // get the data
            Spreadsheet response = request.Execute();

            // print the data
            if (response != null && response.Sheets.Count > 0)
            {
                // should just be one sheet
                return response.Sheets[0];
            }

            // return
            return null;
        }

        public static string PutCell(string sheetId, int row, string col, string value)
        {
            // create value range
            ValueRange inputValue = new ValueRange();
            inputValue.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS
            var oblist = new List<object>() { value };
            inputValue.Values = new List<IList<object>> { oblist };

            // select the cell
            string range = $"{col}{row}";

            // create the request with the sheet ID
            var request = GetService().Spreadsheets.Values.Append(inputValue, sheetId, range);
            request.ValueInputOption = ValueInputOptionEnum.USERENTERED;
            // get the data
            AppendValuesResponse response = request.Execute();

            // print the data
            if (response != null && response.Updates.UpdatedCells > 0)
            {
                // should just be one sheet
                return response.Updates.UpdatedRange;
            }

            // return
            return null;
        }

        public static void PutCell(string sheet, int row, string col, object cell) { }

        public static IList<object> GetRow(string sheetId, int row)
        {
            // select the entire row
            string range = $"{row}:{row}";
            // create the request with the sheet ID
            var request = GetService().Spreadsheets.Values.Get(sheetId, range);

            // get the data
            ValueRange response = request.Execute();

            // print the data
            if (response != null && response.Values.Count > 0)
            {
                // should just be one sheet
                return response.Values[0];
            }

            // return
            return null;
        }

        /// <summary>
        /// writes the sheet in tabular form to standard output
        /// </summary>
        /// <param name="sheet"> A google spreadsheet </param>
        public static void ConsoleLogSheet(Sheet sheet)
        {
            if (sheet != null)
            {
                DataTable table = new DataTable();
                // header
                foreach (var item in sheet.Data[0].RowData[0].Values)
                {
                    table.Columns.Add(item.FormattedValue, typeof(string));
                }

                var body = sheet.Data[0].RowData.Skip(1);
                // rows
                foreach (var rowData in body)
                {
                    table.Rows.Add(rowData.Values.Select(r => r.FormattedValue).ToArray());
                }

                // print
                ConsoleTableBuilder
                .From(table)
                .ExportAndWriteLine();
            }
            else
            {
                Console.WriteLine("No data found");
            }
        }
        public static void ConsoleLogRow(IList<object> row)
        {
            DataTable table = new DataTable();
            // header
            foreach (var item in row)
            {
                table.Columns.Add(item.ToString(), typeof(string));
            }

            // pad with empty row
            table.Rows.Add();

            // print
            ConsoleTableBuilder
            .From(table)
            .ExportAndWriteLine();
        }
    }
}
