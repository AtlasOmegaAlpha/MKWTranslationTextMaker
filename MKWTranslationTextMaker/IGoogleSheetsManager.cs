using Google.Apis.Sheets.v4.Data;

namespace MKWTranslationTextMaker
{
    public interface IGoogleSheetsManager
    {
        Spreadsheet GetSpreadSheet(string googleSpreadsheetIdentifier);

        ValueRange GetSingleValue(string googleSpreadsheetIdentifier, string valueRange);

        BatchGetValuesResponse GetMultipleValues(string googleSpreadsheetIdentifier, string[] ranges);
    }
}
