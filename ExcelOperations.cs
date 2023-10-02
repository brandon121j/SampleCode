using System;
using System.ComponentModel;
using ClosedXML.Excel;

namespace YourNamespace
{
    public interface IExcelOperations : IDisposable
    {
        event ProgressChangedEventHandler? ProgressChanged;
        event Action<object, RunWorkerCompletedEventArgs>? WorkComplete;
        void GenerateReportAsync();
        void Cancel();
    }

    public sealed class ExcelOperations : IExcelOperations
    {
        // Your existing code here

        public event ProgressChangedEventHandler? ProgressChanged;
        public event Action<object, RunWorkerCompletedEventArgs>? WorkComplete;

        public void GenerateReportAsync()
        {
            // Your existing implementation
        }

        public void Cancel()
        {
            // Your existing implementation
        }

        public void Dispose()
        {
            // Your existing implementation
        }
    }
}
