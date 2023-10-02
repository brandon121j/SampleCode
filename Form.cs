using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SF449_Retiree_Generator.Forms
{
    public partial class LoadingForm : Form
    {
        private readonly IExcelOperations _excelOperations;
        // Other fields

        public LoadingForm(IExcelOperations excelOperations, string pnrchkList, string retireeStatement)
        {
            InitializeComponent();
            _excelOperations = excelOperations;
            // Initialize other fields

            _ = RunExcelOperationsAsync(pnrchkList, retireeStatement);
        }

        // Existing methods remain unchanged

        private async Task RunExcelOperationsAsync(string pnrchkList, string retireeStatement)
        {
            _excelOperations.PnrchkListFile = pnrchkList;
            _excelOperations.RetireeStatementsFile = retireeStatement;

            _excelOperations.ProgressChanged += BackgroundWorker_ProgressChanged;
            _excelOperations.WorkComplete += BackgroundWorker_WorkComplete;

            await Task.Run(_excelOperations.GenerateReportAsync);
        }
    }
}

// Define the IExcelOperations interface
public interface IExcelOperations
{
    event ProgressChangedEventHandler ProgressChanged;
    event RunWorkerCompletedEventHandler WorkComplete;
    string PnrchkListFile { get; set; }
    string RetireeStatementsFile { get; set; }
    Task GenerateReportAsync();
}
