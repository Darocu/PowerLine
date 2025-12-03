using System;
using System.Collections.Generic;
using System.Linq;
using TriTech.Common.Interface;
using TriTech.VisiCAD.Adapters;
using TriTech.VisiCAD.App.WPF.Command;
using TriTech.VisiCAD.App.WPF.Plugin;
using TriTech.VisiCAD.Interfaces;

namespace TriTech.Plugin.CUSTPowerLine;

/// <summary>
/// WorkstationPlugin for CUSTPowerLine
/// To enable a PowerLine in CAD you need to go to Toolbox --> Permission Security Manager --> Function Level Security
/// Find the Functionality group you want to enable the PowerLine for, and then find PowerLine - {CommandName} in the list.
/// Double click "No" to change it to "Yes" and save.
/// Go to Configuration Utility --> System Tools --> PowerLine Configuration --> Commands --> Assign Commands
/// Filter for {CommandName} and select "Next". Set the "Description" to the name. Set the "Agency". Set the "Command".
/// Select "Assign" to assign the command. Select "Save" to save the configuration.
///
/// To test this, build the project and take the file TriTech.Plugin.CUSTPowerLine.dll from the output folder and
/// overwrite the existing file in the C:\TriTech\Staging\TriTech\VisiCad\NET folder of your VisiCAD installation.
///
/// To deploy this, build the project and take the file TriTech.Plugin.CUSTPowerLine.dll from the output folder and
/// make a copy of the original TriTech.Plugin.CUSTPowerLine.dll file, archive it, then overwrite the existing
/// file in the test or production file server located in the D:\VisiCAD\TriTech\VisiCAD\NET folder.
/// </summary>

public class WorkstationPlugin : WorkstationPluginBase
{
    internal const string PluginName = "CUSTPowerLine";
    
    private ICADManager CADManagerAdapter { get; set; }

    public override void Start(string pluginConfiguration)
    {
        base.Start(pluginConfiguration);

        CADManagerAdapter = new CADManagerAdapter(CADManager);

        CreateCustomCommandsIfNeeded();
    }
    
    public override void Stop()
    {
        base.Stop();
    }

    private void CreateCustomCommandsIfNeeded()
    {
        var existingCommands = new HashSet<string>();
        var allCustomCommands = CADManager.CommandLineQueryEngine.GetCustomCommands();
        foreach (var commandLineAction in allCustomCommands.Where(commandLineAction => commandLineAction.CustomCommandWorkstationPluginName.Equals(WorkstationPlugin.PluginName, StringComparison.OrdinalIgnoreCase)))
            existingCommands.Add(commandLineAction.CustomCommandName);
        
        // Create the ServiceNow command if needed
        if (!existingCommands.Contains(ServiceNowCommand.CommandName))
        {
            CADManager.CommandLineActionEngine.AddCustomCommand(
                VisiCADDefinition.CommandCategory.Enterprise.ToString(), // ActionType must be a value from CommandCategory - Interface, Enterprise, Unit, Combined
                ServiceNowCommand.CommandName, // CommandDescription
                PluginName,
                ServiceNowCommand.CommandName, //CommandName
                out _
            );
            
        }
        else if (!existingCommands.Contains(NoMctCommand.CommandName))
        {
            CADManager.CommandLineActionEngine.AddCustomCommand(
                VisiCADDefinition.CommandCategory.Enterprise.ToString(), // ActionType must be a value from CommandCategory - Interface, Enterprise, Unit, Combined
                NoMctCommand.CommandName, // CommandDescription
                PluginName,
                NoMctCommand.CommandName, //CommandName
                out _
            );
        }
        else if (!existingCommands.Contains(UpdateUnitStatusLocationCommand.CommandName))
        {
            CADManager.CommandLineActionEngine.AddCustomCommand(
                VisiCADDefinition.CommandCategory.Enterprise.ToString(), // ActionType must be a value from CommandCategory - Interface, Enterprise, Unit, Combined
                UpdateUnitStatusLocationCommand.CommandName, // CommandDescription
                PluginName,
                UpdateUnitStatusLocationCommand.CommandName, //CommandName
                out _
            );
        }
    }
    
    public override CommandBase GetCommand(string commandName)
    {
        CommandBase command;
        switch(commandName)
        {
            case ServiceNowCommand.CommandName:
                command = new ServiceNowCommand(CADManagerAdapter);
                break;
            case NoMctCommand.CommandName:
                command = new NoMctCommand(CADManagerAdapter);
                break;
            case UpdateUnitStatusLocationCommand.CommandName:
                command = new UpdateUnitStatusLocationCommand(CADManagerAdapter);
                break;
            default:
                command = null;
                break;
        }

        return command;
    }
}
