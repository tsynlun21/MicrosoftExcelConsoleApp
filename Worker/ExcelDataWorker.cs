using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MicrosoftExcelServiceConsoleApp.Worker;

public class ExcelDataWorker : IWorkerService
{
    private readonly string _path;

    private IXLTable _productsTable;
    private IXLTable _clientsTable;
    private IXLTable _requestsTable;

    public ExcelDataWorker(string pathToXlsx)
    {
        _path = pathToXlsx;
    }

    #region Public methods
    public void DisplayInformationByProductName(string productName)
    {
        if (string.IsNullOrEmpty(productName))
        {
            Console.WriteLine("Некорректное название товара\n");
            return;
        }
        using var workbook = new XLWorkbook(_path);
        InitializeTables(workbook);


        var productCodeNumber = GetCodeNumberByName(_productsTable, productName);
        if (productCodeNumber is null)
        {
            Console.WriteLine("Товар не найден\n");
            return;
        }
        var productPrice = GetProductPrice(_productsTable, productName);

        var requestsRows = GetRequestsByProductCode(_requestsTable, productCodeNumber);

        if (requestsRows.Count == 0)
        {
            Console.WriteLine("Заявок по выбранному товару не найдено.");
        }
        else
        {
            Console.WriteLine($"{"Организация",-15} | {"Контактное лицо",-30} | {"Количество",-10} | {"Стоимость",-10} | {"Дата заказа",-10}");
            Console.WriteLine();

            foreach (var request in requestsRows)
            {
                var neededAmount = GetValueFromTableRowByField<string>(request, "Требуемое количество");
                var placeDate = GetValueFromTableRowByField<DateTime>(request, "Дата размещения");

                var clientCode = GetValueFromTableRowByField<string>(request, "Код клиента");
                var clientRow = GetFirstRowFromTableByField(_clientsTable, "Код клиента", clientCode);
                var orgName = GetValueFromTableRowByField<string>(clientRow, "Наименование организации");
                var contactPerson = GetValueFromTableRowByField<string>(clientRow, "Контактное лицо (ФИО)");

                Console.WriteLine($"{orgName,-15} | {contactPerson,-30} | {neededAmount,-10} | {int.Parse(productPrice) * int.Parse(neededAmount),-10} | {placeDate.ToShortDateString(),-10}");
            }
        }

        Console.WriteLine();
        Console.ReadKey();
        workbook.Dispose();
    }

    public void ChangeContactPerson(string orgName, string newContactName)
    {
        using var workbook = new XLWorkbook(_path);

        _clientsTable = GetTableByWorksheetName(workbook, "Клиенты");

        var rowToChange = GetFirstRowFromTableByField(_clientsTable, "Наименование организации", orgName.Trim());

        if (rowToChange is null)
        {
            Console.WriteLine("Нет данных о выбранной организации\n");
        }
        else
        {
            var contactField = rowToChange.Field("Контактное лицо (ФИО)");
            var oldContactName = contactField.Value.GetText();

            if (oldContactName == newContactName)
            {
                Console.WriteLine("Новое контактное лицо совпадает со старым. Продолжить? Y/N");

                ConsoleKeyInfo result = new ConsoleKeyInfo();
                while (result.Key != ConsoleKey.Y && result.Key != ConsoleKey.N)
                {
                    result = Console.ReadKey();
                    Console.WriteLine();
                }

                if (result.Key == ConsoleKey.N)
                {
                    return;
                }
            }

            contactField.Value = newContactName.Trim();
            workbook.Save();

            Console.WriteLine($"Контактное лицо \"{orgName}\" изменено с {oldContactName} на {newContactName.Trim()}\n");
            Console.ReadKey();
        }

    }

    public void DisplayGoldenClientByMonthAndYear(int month, int year)
    {
        if (month < 1 || month > 12 || year < 0)
        {
            Console.WriteLine("Некорректный данные");
            return;
        }

        using var workbook = new XLWorkbook(_path);

        _clientsTable = GetTableByWorksheetName(workbook, "Клиенты");
        _requestsTable = GetTableByWorksheetName(workbook, "Заявки");

        var clientCode = _requestsTable.DataRange.Rows().Where(row =>
                row.Field("Дата размещения").GetDateTime().Month == month &&
                row.Field("Дата размещения").GetDateTime().Year == year)
            .OrderByDescending(row => row.Field("Требуемое количество").GetDouble())
            .Select(row => row.Field("Код клиента").GetString()).FirstOrDefault();

        if (clientCode is null)
        {
            Console.WriteLine("Информации о заявках за выбранный период времени не найдено\n");
            return;
        }

        var topCustomer = GetFirstRowFromTableByField(_clientsTable, "Код клиента", clientCode);

        Console.WriteLine($"\"Золотым\" покупателем за {month}.{year} является {topCustomer.Field("Наименование организации").GetString()} (контактное лицо - {topCustomer.Field("Контактное лицо (ФИО)").GetString()})\n");

        Console.ReadKey();

    }

    public bool TryOpenConnection()
    {
        try
        {
            using var workbook = new XLWorkbook(_path);
            return true;
        }
        catch (Exception e)
        {
            return false;
        }
    }
    #endregion

    #region Private methods

    private void InitializeTables(XLWorkbook workbook)
    {
        _productsTable = GetTableByWorksheetName(workbook, "Товары");
        _clientsTable = GetTableByWorksheetName(workbook, "Клиенты");
        _requestsTable = GetTableByWorksheetName(workbook, "Заявки");
    }

    private IXLTable GetTableByWorksheetName(XLWorkbook wb, string sheetName)
    {
        var sheet = wb.Worksheet(sheetName);

        return sheet.Range(sheet.FirstCellUsed(), sheet.LastCellUsed()).AsTable();

    }
    private string? GetCodeNumberByName(IXLTable productsTable, string productName)
    {
        var productRow = GetFirstRowFromTableByField(productsTable, "Наименование", productName);

        return productRow?.Field("Код товара").GetValue<string>() ;

    }

    private string GetProductPrice(IXLTable productsTable, string productName)
    {
        var productRow = GetFirstRowFromTableByField(productsTable, "Наименование", productName);

        return productRow.Field("Цена товара за единицу").GetValue<string>();

    }

    private IXLTableRow? GetFirstRowFromTableByField(IXLTable table, string fieldName, string fieldValue)
    {
        return table.DataRange.Rows()
            .FirstOrDefault(productsRow => productsRow.Field(fieldName).GetString() == fieldValue);
    }

    private List<IXLTableRow> GetRequestsByProductCode(IXLTable requestsTable, string productCodeNumber)
    {
        return requestsTable.DataRange.Rows()
            .Where(row => row.Field("Код товара").GetString() == productCodeNumber).ToList();
    }

    private T GetValueFromTableRowByField<T>(IXLTableRow row, string field)
    {
        return row.Field(field).GetValue<T>();
    }
    #endregion
}