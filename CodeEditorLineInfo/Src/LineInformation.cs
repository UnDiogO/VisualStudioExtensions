using System;
using System.ComponentModel.Design;

using EnvDTE;
using EnvDTE80;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace CodeEditorLineInfo.Src
{
    internal sealed class LineInformation
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("4db00d78-12d0-4e91-b5cd-2571752fe8eb");

        private readonly AsyncPackage package;

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return package;
            }
        }

        public static LineInformation Instance
        {
            get;
            private set;
        }

        private LineInformation(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new LineInformation(package, commandService);
        }

        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var serviceTextManager = await ServiceProvider.GetServiceAsync(typeof(SVsTextManager));
            var textManager = serviceTextManager as IVsTextManager2;
            textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out IVsTextView textView);
            textView.GetCaretPos(out int lineIndex, out int columnIndex);
            textView.GetBuffer(out IVsTextLines buffer);
            buffer.GetLengthOfLine(lineIndex, out int lineLength);
            buffer.GetLineText(lineIndex, 0, lineIndex, lineLength, out string lineText);

            var dteService = await ServiceProvider.GetServiceAsync(typeof(DTE));
            var dte = dteService as DTE2;
            var activeDocument = dte.ActiveDocument;

            var resultText = $"{activeDocument.FullName} at line {lineIndex + 1}{Environment.NewLine}{lineText.TrimStart(' ')}";

            System.Windows.Forms.Clipboard.SetText(resultText);

            var serviceBar = await ServiceProvider.GetServiceAsync(typeof(SVsStatusbar));
            var statusBar = serviceBar as IVsStatusbar;
            statusBar.SetText("Line information copied to clipboard");
        }
    }
}
