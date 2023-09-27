public static class ConsoleViewer
{
    public static void EnterPath() =>
        Console.WriteLine("Введите полный путь до файла .xlsx");

    public static void ShowMenu()
    {
        Console.WriteLine("Выберите пункт меню:" +
                          "\n\t1. Вывести информацию о клиентах, заказавших определенный товар." +
                          "\n\t2. Изменить контактное лицо организации." +
                          "\n\t3. Определение \"золотого\" клиента в указаном месяце и году" +
                          "\n\t0. Завершить работу");
    }

    public static void IncorrectPath() =>
        Console.WriteLine("\nНекорректный путь. Перепроверьте введенные данные и повторите попытку.");

    public static void EnterProductName() =>
        Console.Write("Введите название продукта -> ");

    public static void EnterOrganizationName() =>
        Console.Write("Введите наименование организации -> ");

    public static void EnterNewContactPerson() =>
        Console.Write("\nВведите новое контактное лицо -> ");

    public static void EnterYear() =>
        Console.Write("Введите год -> ");

    public static void EnterMonth() =>
        Console.Write("Введите месяц -> ");

    public static void IncorrectData() =>
        Console.WriteLine("Некорректные данные. Перепроверьте введенные данные и повторите попытку.");
}

