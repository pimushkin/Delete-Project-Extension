using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using lab5Extension.Properties;
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
            if (activeProjects.Length == 0)
            {
                VsShellUtilities.ShowMessageBox(
                    package,
                    Resources.MissingProjectErrorMessage,
                    null,
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }
            var message = activeProjects.Length == 1
                ? string.Format(Resources.DeletingProjectMessage, $"'{activeProjects.First().Name}'")
                : string.Format(Resources.DeletingProjectsMessage, string.Join(", ", activeProjects.Select(x => $"'{x.Name}'")));

            // Show a message box to prove we were here
            var messageBoxResult = VsShellUtilities.ShowMessageBox(
                package,
                message,
                null,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            var exceptions = new Dictionary<dynamic, Exception>();
            if (messageBoxResult != 1) return;
            {
                var error = false;
                var solutionPath = Path.GetDirectoryName(projectInterface.Solution.FileName)?.TrimEnd('\\');
                foreach (var project in activeProjects)
                {
                    try
                    {
                        string projectPath = Path.GetDirectoryName(project.FileName)?.TrimEnd('\\');
                        if (string.Equals(solutionPath, projectPath, StringComparison.OrdinalIgnoreCase))
                        {
                            exceptions.Add(project, new Exception(Resources.SameDirectoryErrorMessage));
                        }
                        else
                        {
                            projectInterface.Solution.Remove(project);
                            FileSystem.DeleteDirectory(projectPath, UIOption.OnlyErrorDialogs,
                                RecycleOption.SendToRecycleBin, UICancelOption.DoNothing);
                        }
                    }
                    catch (Exception exception)
                    {
                        error = true;
                        exceptions.Add(project, exception);
                    }
                }
                if (!exceptions.Any()) return;
                var resultMessage = string.Join(Environment.NewLine, exceptions.Select(x => $"'{x.Key.Name}': {x.Value.Message}"));
                VsShellUtilities.ShowMessageBox(
                    package,
                    resultMessage,
                    null,
                    error ? OLEMSGICON.OLEMSGICON_CRITICAL : OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }
    }
}
