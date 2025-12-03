using System;
using System.Collections.Generic;
using System.Linq;
using TriTech.Common.Interface;
using TriTech.VisiCAD.App.WPF.Command;
using TriTech.VisiCAD.App.WPF.Properties;
using TriTech.VisiCAD.Interfaces;

namespace TriTech.Plugin.CUSTPowerLine
{
    public class UpdateUnitStatusLocationCommand : CommandBase
    {
        internal const string CommandName = "Update Unit Status and Location (ONS/TC)";
        private const string ParameterComment = "Comment";

        private List<string> _units;
        private string _comment;
        
        public UpdateUnitStatusLocationCommand(ICADManager cadManager) : base(cadManager)
        {
            Target = new BasicCommandTarget(Resources.UnitListPrompt, CommandTargetType.UnitList, true);
            
            m_parameters.Add(new BasicCommandParameter(
                CommandParameterType.NotAvailable,
                ParameterComment,
                true));
        }

        public override ICommandTarget Target { get; }

        public override void Execute()
        {
            var result = new BasicCommandResult(CommandState.Pending, this);

            try
            {
                _units = Target.Values?.ToList();
                _comment = Parameters.FirstOrDefault(p => p.Name == ParameterComment)?.Value ?? string.Empty;

                if (_units == null) return;
                CADManager.GeneralActionEngine.AddActivityLogEntry("Update Unit Location",
                    $"Executing {CommandName} Command for units: [{string.Join(", ", _units)}]");

                foreach (var unit in _units)
                {
                    var unitInfo = CADManager.UnitQueryEngine.GetUnitByName(unit);
                    if (unitInfo == null)
                    {
                        return;
                    }

                    var parameters = new List<string>()
                    {
                        _comment,
                    };

                    if (unitInfo.Status is VisiCADDefinition.EnumStatus.DepartScene)
                    {
                        var commandEngine = new CommandEngine(CADManager, "clsAction", "ActivateUnitAtHospital",
                            unit, parameters, this);
                        commandEngine.ExecuteCommand(CommandEngine_UpdateStatus);
                    }
                    else
                    {
                        var commandEngine = new CommandEngine(CADManager, "clsAction", "ActivateUnitArivalAtScene",
                            unit, parameters, this);
                        commandEngine.ExecuteCommand(CommandEngine_UpdateStatus);
                    }
                }

                CADManager.GeneralActionEngine.AddActivityLogEntry("Update Unit Location",
                    $"Completed {CommandName} Command for units: [{string.Join(", ", _units)}]");
            }
            catch (Exception ex)
            {
                HandleCommandException(ex);
            }
        }

        private void CommandEngine_UpdateStatus(object sender, CommandResultEventArgs e)
        {
            if (e.CommandResult.State == CommandState.Failure)
            {
                m_commandResult.State = CommandState.Failure;
                m_commandResult.Validations.Add(new CommandValidation(MessageLevel.Error,
                    e.CommandResult.GetFailureMessage(), OriginalSourceText));
            }
            else
            {
                m_commandResult.State = CommandState.Success;
            }

            InvokeCommandComplete(m_commandResult);
        }

        private void HandleCommandException(Exception ex)
        {
            HandleException(ex);
            m_commandResult.Validations.Add(new CommandValidation(MessageLevel.Error,
                $"Failed to Execute {CommandName} Command: {ex.Message}", string.IsNullOrWhiteSpace(this.UserEnteredText) ? string.Empty : this.UserEnteredText));
            m_commandResult.State = CommandState.Failure;
            InvokeCommandComplete(m_commandResult);
        }
    }
}