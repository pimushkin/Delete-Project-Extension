using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace lab5Extension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5d66f7d9-3236-40f6-8513-623e60fa5424");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var projectInterface = (DTE)await ServiceProvider.GetServiceAsync(typeof(DTE));
            Assumes.Present(projectInterface);
            var activeProjects = (dynamic[])projectInterface.ActiveSolutionProjects;
            const string message = "The project(s) and all content will be deleted from the computer.";

            // Show a message box to prove we were here
            var messageBoxResult = VsShellUtilities.ShowMessageBox(
                package,
                message,
                null,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            if (messageBoxResult != 1) return;
            {
                var solutionPath = Path.GetDirectoryName(projectInterface.Solution.FileName)?.TrimEnd('\\');
                foreach (var project in activeProjects)
                {
                    try
                    {
                        string projectPath = Path.GetDirectoryName(project.FileName)?.TrimEnd('\\');
                        if (string.Equals(solutionPath, projectPath, StringComparison.OrdinalIgnoreCase))
                        {
                            VsShellUtilities.ShowMessageBox(
                                package,
                                "An error occurred during deletion because the project directory matches the solution directory.",
                                null,
                                OLEMSGICON.OLEMSGICON_WARNING,
                                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                            continue;
                        }
                        projectInterface.Solution.Remove(project);
                        FileSystem.DeleteDirectory(projectPath, UIOption.OnlyErrorDialogs,
                            RecycleOption.SendToRecycleBin, UICancelOption.DoNothing);
                    }
                    catch (Exception exception)
                    {
                        VsShellUtilities.ShowMessageBox(
                            package,
                            exception.Message,
                            null,
                            OLEMSGICON.OLEMSGICON_CRITICAL,
                            OLEMSGBUTTON.OLEMSGBUTTON_OK,
                            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    }
                }
            }
        }
    }
}
