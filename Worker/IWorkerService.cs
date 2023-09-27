namespace MicrosoftExcelServiceConsoleApp.Worker;

public interface IWorkerService
{
    public void DisplayInformationByProductName(string productName);

    public void ChangeContactPerson(string orgName, string newContactName);

    public void DisplayGoldenClientByMonthAndYear(int month,  int year);
}