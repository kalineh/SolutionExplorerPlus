using System;
using System.IO;
using System.Collections.Generic;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.VCProjectEngine;

namespace SolutionExplorerPlus
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect
        : IDTExtensibility2
        , IDTCommandTarget
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;

            var commands = _applicationObject.Commands as Commands2;
            var command_bars = _applicationObject.CommandBars as CommandBars;

            var solution_bar = command_bars["Solution"] as CommandBar;
            var project_bar = command_bars["Project"] as CommandBar;
            var folder_bar = command_bars["Folder"] as CommandBar;
            var item_bar = command_bars["Item"] as CommandBar;

            // Solution Folder: right-click solution explorer top-level icon
            // Project: right-click project in solution explorer
            // Folder: right-click folder in solution explorer
            // Item: right-click item in solution explorer

            var repopulate_solution = commands.AddNamedCommand(_addInInstance, "RepopulateFromSolution", "Repopulate", null, false);
            var repopulate_project = commands.AddNamedCommand(_addInInstance, "RepopulateFromProject", "Repopulate", null, false);
            var repopulate_folder = commands.AddNamedCommand(_addInInstance, "RepopulateFromFolder", "Repopulate", null, false);
            var repopulate_item = commands.AddNamedCommand(_addInInstance, "RepopulateFromItem", "Repopulate", null, false);

            repopulate_solution.AddControl(solution_bar);
            repopulate_project.AddControl(project_bar);
            repopulate_folder.AddControl(folder_bar);
            repopulate_item.AddControl(item_bar);
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
            var commands = _applicationObject.Commands as Commands2;
            var to_delete = new List<Command>();

            foreach (Command command in commands)
            {
                if (command.Name.StartsWith(_addInInstance.ProgID))
                {
                    to_delete.Add(command);
                }
            }

            foreach (Command command in to_delete)
            {
                command.Delete();
            }
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
        //
        // IDTCommandTarget
        //

		/// <summary></summary>
        public void Exec(string CmdName, vsCommandExecOption ExecuteOption, ref object VariantIn, ref object VariantOut, ref bool Handled)
        {
            _applicationObject.ToolWindows.OutputWindow.ActivePane.OutputString(String.Format("Exec:{0}\n", CmdName));

            switch (CmdName)
            {
                case "SolutionExplorerPlus.Connect.RepopulateFromSolution": RepopulateFromSolution(); break;
                case "SolutionExplorerPlus.Connect.RepopulateFromProject": RepopulateFromProject(); break;
                case "SolutionExplorerPlus.Connect.RepopulateFromFolder": RepopulateFromFolder(); break;
                case "SolutionExplorerPlus.Connect.RepopulateFromItem": RepopulateFromItem(); break;
            }
        }

		/// <summary></summary>
        public void QueryStatus(string CmdName, vsCommandStatusTextWanted NeededText, ref vsCommandStatus StatusOption, ref object CommandText)
        {
            StatusOption = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
            CommandText = CmdName.Substring(_addInInstance.ProgID.Length + 1);
        }

        //
        // Actual plugin functionality
        //

        private void Log(string s)
        {
            _applicationObject.ToolWindows.OutputWindow.ActivePane.OutputString(String.Format("{0}\n", s));
        }

        private struct QueueEntry
        {
            public VCFilter _vcfilter;
            public VCFile _vcfile;
        };

        private void QueueProject(List<QueueEntry> queue, VCProject src)
        {
            foreach (var item in src.Items)
            {
                var vcfilter = item as VCFilter;
                if (vcfilter != null)
                {
                    QueueFiltersRecursive(queue, vcfilter);
                }

                var vcfile = item as VCFile;
                if (vcfile != null && vcfile.Parent as VCFilter != null)
                {
                    queue.Add(new QueueEntry() { _vcfile = vcfile, _vcfilter = vcfile.Parent as VCFilter });
                }
            }
        }

        private void QueueFilter(List<QueueEntry> queue, VCFilter src)
        {
            foreach (var item in src.Items)
            {
                var vcfile = item as VCFile;
                if (vcfile != null && vcfile.Parent as VCFilter != null)
                {
                    queue.Add(new QueueEntry() { _vcfile = vcfile, _vcfilter = vcfile.Parent as VCFilter });
                }
            }
        }

        private void QueueFiltersRecursive(List<QueueEntry> queue, VCFilter src)
        {
            foreach (var item in src.Items)
            {
                var vcfilter = item as VCFilter;
                if (vcfilter != null)
                {
                    QueueFiltersRecursive(queue, vcfilter);
                }

                var vcfile = item as VCFile;
                if (vcfile != null && vcfile.Parent as VCFilter != null)
                {
                    queue.Add(new QueueEntry() { _vcfile = vcfile, _vcfilter = vcfile.Parent as VCFilter });
                }
            }
        }

        private void QueueFile(List<QueueEntry> queue, VCFile src)
        {
            if (src.Parent as VCFilter != null)
            {
                queue.Add(new QueueEntry() { _vcfile = src, _vcfilter = src.Parent as VCFilter });
            }
        }

        private void QueueStripDuplicates(List<QueueEntry> queue)
        {
            var uniques = new List<QueueEntry>();
            var seen = new Dictionary<VCFilter, List<string>>();

            foreach (var lhs in queue)
            {
                var vcfile = lhs._vcfile;
                var vcfilter = lhs._vcfilter;
                var extension = vcfile.Extension;

                // add any filters we havent seen before
                if (!seen.ContainsKey(vcfilter))
                {
                    uniques.Add(new QueueEntry() { _vcfile = vcfile, _vcfilter = vcfilter });
                    seen.Add(vcfilter, new List<string>() { extension, });
                    continue;
                }

                // check if this extension has been seen for this filter
                var exists = false;
                var seen_extensions = seen[vcfilter];
                foreach (var e in seen_extensions)
                {
                    if (e != extension)
                        continue;

                    exists = true;
                    break;
                }

                if (!exists)
                {
                    uniques.Add(new QueueEntry() { _vcfile = vcfile, _vcfilter = vcfilter });
                    seen_extensions.Add(extension);
                }
            }

            queue.Clear();
            queue.AddRange(uniques);
        }

        private void QueueProcess(List<QueueEntry> queue)
        {
            QueueStripDuplicates(queue);

            foreach (var item in queue)
            {
                var vcfile = item._vcfile;
                var vcfilter = item._vcfilter;

                string fullpath = vcfile.FullPath;
                string extension = vcfile.Extension;
                string folder = Path.GetDirectoryName(fullpath);

                var search = String.Format("*{0}", extension);
                var siblings = Directory.GetFiles(folder, search, SearchOption.TopDirectoryOnly);

                foreach (var sibling in siblings)
                {
                    if (!vcfilter.CanAddFile(sibling))
                        continue;

                    Log(String.Format("> : adding {0}", sibling));

                    vcfilter.AddFile(sibling);
                }
            }
        }

        private void RepopulateFromSolution()
        {
            var queue = new List<QueueEntry>();
            var projects = _applicationObject.DTE.Solution.Projects;

            foreach (Project project in projects)
            {
                var vcproject = project.Object as VCProject;

                QueueProject(queue, vcproject);
            }

            QueueProcess(queue);
        }

        private void RepopulateFromProject()
        {
            var queue = new List<QueueEntry>();
            var selected = _applicationObject.DTE.SelectedItems;

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning project {0}...", s.Name));

                var vcproject = s.ProjectItem.Object as VCProject;
                if (vcproject == null)
                    continue;

                QueueProject(queue, vcproject);
            }

            QueueProcess(queue);
        }

        private void RepopulateFromFolder()
        {
            var queue = new List<QueueEntry>();
            var selected = _applicationObject.DTE.SelectedItems;

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning folder {0}...", s.Name));

                var vcfilter = s.ProjectItem.Object as VCFilter;
                if (vcfilter == null)
                    continue;

                QueueFilter(queue, vcfilter);
            }

            QueueProcess(queue);
        }

        private void RepopulateFromItem()
        {
            var queue = new List<QueueEntry>();
            var selected = _applicationObject.DTE.SelectedItems;

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning item {0}...", s.Name));

                var vcfile = s.ProjectItem.Object as VCFile;

                if (vcfile == null)
                    continue;

                QueueFile(queue, vcfile);
            }

            QueueProcess(queue);
        }


        /*
        private void RepopulateSiblings(VCFile vcfile)
        {
            string fullpath = vcfile.FullPath;
            string extension = vcfile.Extension;
            string folder = Path.GetDirectoryName(fullpath);

            Log(String.Format("> VCFile: repopulating type {0} at {1}", extension, folder));

            var vcfilter = vcfile.Parent as VCFilter;
            if (vcfilter != null)
            {
                var vcproject = vcfile.project as VCProject;
                var search = String.Format("*{0}", extension);
                var siblings = Directory.GetFiles(folder, search, SearchOption.TopDirectoryOnly);

                foreach (var sibling in siblings)
                {
                    if (!vcfilter.CanAddFile(sibling))
                        continue;

                    Log(String.Format("> : adding {0}", sibling));

                    vcfilter.AddFile(sibling);
                }
            }
        }
        */
        
        /*
        private void RepopulateFromSolution()
        {
            var dte = _applicationObject.DTE;
            var projects = _applicationObject.DTE.ActiveSolutionProjects as Array;
            var selected = _applicationObject.DTE.SelectedItems;

            var solution = _applicationObject.DTE.Solution;

            foreach (Project project in solution.Projects)
            {
                var vcproject = project.Object as VCProject;

                RepopulateProject(vcproject);
            }
        }

        private void RepopulateFromProject()
        {
            var projects = _applicationObject.DTE.ActiveSolutionProjects as Array;
            var selected = _applicationObject.DTE.SelectedItems;

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning project {0}...", s.Name));

                var vcproject = s.ProjectItem.Object as VCProject;
                if (vcproject == null)

                    continue;

                RepopulateProject(vcproject);
            }
        }

        private void RepopulateFromFolder()
        {
            var projects = _applicationObject.DTE.ActiveSolutionProjects as Array;
            var selected = _applicationObject.DTE.SelectedItems;

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning folder {0}...", s.Name));

                var vcfilter = s.ProjectItem.Object as VCFilter;
                if (vcfilter == null)
                    continue;

                foreach (var item in vcfilter.Files)
                {
                    var vcfile = item as VCFile;

                    if (vcfile == null)
                        continue;

                    RepopulateSiblings(vcfile);
                }
            }
        }

        private void RepopulateFromItem()
        {
            var projects = _applicationObject.DTE.ActiveSolutionProjects as Array;
            var selected = _applicationObject.DTE.SelectedItems;
            var process = new List<SelectedItem>();

            foreach (SelectedItem s in selected)
            {
                Log(String.Format("Scanning item {0}...", s.Name));

                var vcfile = s.ProjectItem.Object as VCFile;

                if (vcfile == null)
                    continue;

                RepopulateSiblings(vcfile);
            }
        }
         * */

		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}
