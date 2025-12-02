using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using TriTech.Common.Interface;
using TriTech.VisiCAD.App.WPF.Command;
using TriTech.VisiCAD.Incidents;
using TriTech.VisiCAD.Interfaces;
using CUSTPowerLineDialog;

namespace TriTech.Plugin.CUSTPowerLine;

public class NoMctCommand : CommandBase
{
    internal const string CommandName = "MCT Not Available";
    private const string ParameterIncidentNumber = "Incident ID or Number";
    
    private int? _incidentId;
    private string _description;
    private Incident _activeIncident;
    
    public NoMctCommand(ICADManager cadManager) : base(cadManager)
    {
        m_parameters.Add( 
            new BasicCommandParameter(
                CommandParameterType.NotAvailable,
                ParameterIncidentNumber,
                false)
        );
    }
    
    public override ICommandTarget Target => null;

    public override async void Execute()
    {
        try
        {
            var unused = new BasicCommandResult(CommandState.Pending, this);
        
            await CaptureParametersAsync();
            
            if (_incidentId == null || string.IsNullOrEmpty(_description))
                throw new Exception("NoMCT: Invalid incident number or description provided");
            
            CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Executing {CommandName} Command");
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));

            await ExecuteNoMctCommandAsync();
            CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Completed {CommandName} Command for Incident {_incidentId}");
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
        }
        catch (Exception ex)
        {
            HandleCommandException(ex);
        }
    }
    
    private Task CaptureParametersAsync()
    {
        var incidentValue = Parameters.FirstOrDefault(p => p.Name == ParameterIncidentNumber)?.Value;
        if (!string.IsNullOrEmpty(incidentValue))
        {
            while (incidentValue.StartsWith(".") || incidentValue.StartsWith("~"))
                incidentValue = incidentValue.Substring(1);

            try
            {
                if (int.TryParse(incidentValue, out _) && incidentValue.Length == 3)
                {
                    var incidentList = CADManager.IncidentQueryEngine?.GetActiveIncidentIDListByShortcutID(incidentValue);
                    if (incidentList != null) 
                        _incidentId = incidentList.LastOrDefault();
                }
                else if (incidentValue.Length >= 15)
                {
                    _incidentId = CADManager.IncidentQueryEngine.GetIncidentIDByIncidentNumber(incidentValue);
                }
                else
                {
                    var incident = CADManager.IncidentQueryEngine.GetIncident(Convert.ToInt32(incidentValue));
                    
                    if (incident == null)
                    {
                        CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Incident not found for value '{incidentValue}'.");
                        throw new Exception($"Incident not found for value '{incidentValue}'.");
                    }
                    
                    _incidentId = incident.ID;
                }
                
                if (_incidentId == null)
                {
                    CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", "Incident ID is null or invalid.");
                    throw new Exception("Incident ID is null or invalid.");
                }
                
                _activeIncident = CADManager.IncidentQueryEngine?.GetActiveIncident(_incidentId.Value);
                if (_activeIncident == null)
                {
                    CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"No active incident found for Incident ID {_incidentId}.");
                    throw new Exception($"No active incident found for Incident ID {_incidentId}.");
                }
            }
            catch (Exception ex)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Error parsing incident value '{incidentValue}': {ex.Message}");
                throw;
            }
        }
        else
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", "Incident number parameter is missing or empty.");
            throw new Exception("Incident number parameter is missing or empty.");
        }

        string selectedReason = null;
        Application.Current.Dispatcher.Invoke(() =>
        {
            var dialog = new NoMctPowerLineDialog();
            var result = dialog.ShowDialog();
            if (result == true && !string.IsNullOrWhiteSpace(dialog.SelectedReason))
            {
                selectedReason = dialog.SelectedReason;
            }
        });

        if (string.IsNullOrWhiteSpace(selectedReason))
            throw new Exception("No reason selected for MCT unavailability.");

        _description = "NOMCT: " + selectedReason;
        
        return Task.CompletedTask;
    }

    private async Task ExecuteNoMctCommandAsync()
    {
        const VisiCADDefinition.CommentCategory commentCategory = VisiCADDefinition.CommentCategory.PowerlineUserEntered;
        
        try
        {
            var upperDescription = _description?.ToUpperInvariant();

            CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Adding comment for Incident ID: {_incidentId}, Description: {_description}");

            await Task.Run(() => CADManager.IncidentActionEngine.AddIncidentComment(
                _activeIncident.ID,
                CADManager.Personnel.Code,
                upperDescription,
                DateTime.Now,
                commentCategory));
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Error adding incident comment: {ex.Message}");
            throw;
        }
    }
    
    private void HandleCommandException(Exception ex)
    {
        HandleException(ex);
        m_commandResult.Validations.Add(new CommandValidation(MessageLevel.Error,
            $"Failed to Execute {CommandName} Command: {ex.Message}", string.IsNullOrWhiteSpace(this.UserEnteredText) ? string.Empty : this.UserEnteredText));
        CADManager.GeneralActionEngine.AddActivityLogEntry("NoMCT", $"Failed to Execute {CommandName} Command: {ex.Message}");
        m_commandResult.State = CommandState.Failure;
        InvokeCommandComplete(m_commandResult);
    }
}