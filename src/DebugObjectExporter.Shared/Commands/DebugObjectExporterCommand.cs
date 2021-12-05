﻿using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Dynamic;
using System.IO;
using System.Linq;
using Task = System.Threading.Tasks.Task;

namespace DebugObjectExporter.Shared.Commands
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DebugObjectExportCommand
    {
        private string[] SimpleTypes = {
            "bool",
            "bool?",
            "byte",
            "byte?",
            "sbyte",
            "sbyte?",
            "char",
            "char?",
            "decimal",
            "decimal?",
            "double",
            "double?",
            "float",
            "float?",
            "int",
            "int?",
            "uint",
            "uint?",
            "long",
            "long?",
            "ulong",
            "ulong?",
            "object",
            "object?",
            "short",
            "short?",
            "ushort",
            "ushort?",
            "string",
            "System.Guid",
            "System.Guid?"
        };
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("21a5ef58-2019-4e78-9b43-1b70f7943e07");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DebugObjectExportCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DebugObjectExportCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static DebugObjectExportCommand Instance
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
            // Switch to the main thread - the call to AddCommand in TestToolWindowCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new DebugObjectExportCommand(package, commandService);
        }

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            this.package.JoinableTaskFactory.RunAsync(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var dte2 = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE2;
                var expressions = dte2?.Debugger?.CurrentStackFrame?.Locals?
                    .Cast<Expression>()
                    .ToList();
                try
                {
                    if (expressions != null && expressions.Any())
                    {
                        dynamic obj = new ExpandoObject();
                        Process(obj, expressions);
                        var settings = new JsonSerializerSettings { ReferenceLoopHandling = ReferenceLoopHandling.Ignore, Formatting = Formatting.Indented };
                        string result = JsonConvert.SerializeObject(obj, settings);

                        File.WriteAllText("d:/data.json", result);
                    }
                }
                catch (Exception ex)
                {
                    var m = ex.Message;
                }
            });
        }

        private void Process(ExpandoObject storage, List<Expression> expressions)
        {
            foreach (Expression expression in expressions)
            {
                if (expression.IsValidValue)
                {
                    var storageDic = storage as IDictionary<string, object>;
                    if (SimpleTypes.Contains(expression.Type))
                    {
                        storageDic.Add(expression.Name, GetValue(expression.Type, expression.Value));
                        continue;
                    }

                    var members = expression.DataMembers.Cast<Expression>().ToList();
                    if (members.Any())
                    {
                        storageDic.Add(expression.Name, new ExpandoObject());
                        Process((ExpandoObject)storageDic[expression.Name], members);
                    }
                }
            }
        }

        private object GetValue(string type, string value)
        {
            switch (type)
            {
                case "int":
                    return Convert.ToInt32(value);
                case "string":
                    return value.Trim('\"');
                case "System.Guid":
                    return value.Trim('{', '}');
                default:
                    return "";
            }
        }
    }
}
