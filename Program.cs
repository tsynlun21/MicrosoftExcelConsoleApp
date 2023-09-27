using MicrosoftExcelServiceConsoleApp.Worker;
namespace MicrosoftExcelServiceConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ConsoleViewer.EnterPath();
            var path = Console.ReadLine();

            var worker = new ExcelDataWorker(path);

            while (worker.TryOpenConnection() != true)
            {
                ConsoleViewer.IncorrectPath();
                path = Console.ReadLine();
                worker = new ExcelDataWorker(path);
            }

            ConsoleViewer.ShowMenu();
            ConsoleKeyInfo result = Console.ReadKey();
            Console.WriteLine();
            while (result.Key != ConsoleKey.D0)
            {
                switch (result.Key)
                {
                    case ConsoleKey.D1:
                        {
                            ConsoleViewer.EnterProductName();
                            var productName = Console.ReadLine();
                            Console.WriteLine();
                            worker.DisplayInformationByProductName(productName);
                            ConsoleViewer.ShowMenu();
                            result = Console.ReadKey();
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.D2:
                        {
                            ConsoleViewer.EnterOrganizationName();
                            var orgName = Console.ReadLine();
                            ConsoleViewer.EnterNewContactPerson();
                            var newContactPerson = Console.ReadLine();
                            Console.WriteLine();

                            worker.ChangeContactPerson(orgName, newContactPerson);
                            ConsoleViewer.ShowMenu();
                            result = Console.ReadKey();
                            Console.WriteLine();
                            break;
                        }
                    case ConsoleKey.D3:
                        {
                            int year, month;
                            ConsoleViewer.EnterYear();
                            while (!int.TryParse(Console.ReadLine(), out year))
                            {
                                ConsoleViewer.IncorrectData();
                            }

                            ConsoleViewer.EnterMonth();
                            while (!int.TryParse(Console.ReadLine(), out month))
                            {
                                ConsoleViewer.IncorrectData();
                            }
                            Console.WriteLine();

                            worker.DisplayGoldenClientByMonthAndYear(month, year);
                            ConsoleViewer.ShowMenu();
                            result = Console.ReadKey();
                            Console.WriteLine();
                            break;
                        }
                }
            }

        }
    }
}