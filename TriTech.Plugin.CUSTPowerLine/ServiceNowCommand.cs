using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TriTech.VisiCAD.App.WPF.Behaviors;
using TriTech.VisiCAD.App.WPF.Command;
using TriTech.VisiCAD.Interfaces;

namespace TriTech.Plugin.CUSTPowerLine;

public class ServiceNowCommand : CommandBase
{
    // ReSharper disable once InconsistentNaming
    internal const string CommandName = "Enter ServiceNow Ticket";
    private const string ParameterIncidentNumber = "Incident ID or Number";
    private const string ParameterDescription = "Problem Description";

    private const string UserDepartment = "a44ebc89db95b6001f5074b5ae9619c6";
    private const string UserLocation = "Emergency Communication Center";
    private const string StreetAddress = "2000 Radcliff Dr";
    private const string ContactType = "Self-Service";
    private const string ConfigurationItem = "648f5349dbd283c09be3dee5ce9619e9";
    private const string AssignmentGroup = "2079637987dc3010175d1fcd3fbb35bc";
    private const int State = 3; // 1 = new, 3 = on hold
    private const int Impact = 3;
    private const int Urgency = 3;
    private const string UserUrl = "https://cincinnatioh.service-now.com/api/now/table/sys_user";
    private const string IncidentUrl = "https://cincinnatioh.service-now.com/api/now/table/incident";
    private const string FileUploadUrl = "https://cincinnatioh.service-now.com/api/now/attachment/upload";
    
    private readonly string _apiKey = Environment.GetEnvironmentVariable("SERVICENOW_API_KEY");
    
    private string _timeStamp;
    private string _incidentNumber;
    private string _description;
    private string _hostName;
    private int _cadUserId;
    private string _cadChrisId;
    private string _cadName;
    private string _cadWorkstationId;
    private string _snowScreenshotPath;
    private string _snowActivityLogPath;
    private string _snowLogEntriesPath;
    private string _snowIncidentValue;
    private string _snowUserSysId;
    private string _emailSubject;
    private string _emailBody;
    
    public ServiceNowCommand(ICADManager cadManager) : base(cadManager)
    {
        //_target = new BasicCommandTarget(ParameterIncidentNumber, CommandTargetType.CombinedSingleIncident, true);
        
        m_parameters.Add( 
            new BasicCommandParameter(
                CommandParameterType.NotAvailable,
                ParameterIncidentNumber,
                true) // Optional
            );
        m_parameters.Add(
            new BasicCommandParameter(
                CommandParameterType.NotAvailable,
                ParameterDescription,
                false) // Required
        );
    }
    
    public override ICommandTarget Target => null;

    public override async void Execute()
    {
        BasicCommandResult unused = new BasicCommandResult(CommandState.Pending, this);

        try
        {
            if (_apiKey == null)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", "API key is null");
                InvokeCommandComplete(new BasicCommandResult(CommandState.Failure, this));
                return;
            }
            
            await CaptureParameters();
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Executing {CommandName} Command");
            
            if (_cadUserId == 0)
            {
                InvokeCommandComplete(new BasicCommandResult(CommandState.Failure, this));
                return;
            }
            
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"CAD User ID: {_cadUserId}");
            
            await CaptureScreens();
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Screenshot captured: {_snowScreenshotPath}");
            
            await GetSnowUser();
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"ServiceNow User ID: {_snowUserSysId}");
            _snowUserSysId ??= "bcbdc55ddb9836009be3dee5ce961924"; // Default user if not found
            
            await CreateSnowTicket(); // Snow Docs: https://docs.servicenow.com/bundle/washingtondc-api-reference/page/integrate/inbound-rest/concept/c_TableAPI.html
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
            
            // TODO: TEST THIS
            if (_snowIncidentValue == null)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", "Failed to create ServiceNow incident");
                InvokeCommandComplete(new BasicCommandResult(CommandState.Failure, this));
                return;
            }
            // TODO: TEST THIS
            
            if (_snowIncidentValue != null)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"ServiceNow Incident ID: {_snowIncidentValue}");
                
                await UploadFileToServiceNow(_snowScreenshotPath, "incident", _snowIncidentValue);
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Screenshot uploaded: {_snowScreenshotPath}");
                
                await GetLogEntry();
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Log entry captured: {_snowLogEntriesPath}");
                
                await UploadFileToServiceNow(_snowLogEntriesPath, "incident", _snowIncidentValue);
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Log entry uploaded: {_snowLogEntriesPath}");
                
                await GetActivityLogs();
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Activity log captured: {_snowActivityLogPath}");
                
                await UploadFileToServiceNow(_snowActivityLogPath, "incident", _snowIncidentValue);
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Activity log uploaded: {_snowActivityLogPath}");

                await CleanUpGeneratedFiles();
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"File Cleanup Completed");
                
                await SendSupervisorEmail();
                InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Supervisor Email Sent");
            }
            
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Completed {CommandName} Command");
            InvokeCommandComplete(new BasicCommandResult(CommandState.Success, this));
        }
        catch (Exception ex)
        {
            HandleCommandException(ex);
        }
    }

    private Task CaptureParameters()
    {
        try
        {
            // Command parameters
            var incidentValue = Parameters.FirstOrDefault(p => p.Name == ParameterIncidentNumber)?.Value;

            if (!string.IsNullOrEmpty(incidentValue))
            {
                // Remove leading periods and tildes
                while (incidentValue.StartsWith(".") || incidentValue.StartsWith("~"))
                    incidentValue = incidentValue.Substring(1);

                // Check if the incident value is a valid incident number
                if (int.TryParse(incidentValue, out _) && incidentValue.Length == 3)
                {
                    var incidentList = CADManager.IncidentQueryEngine.GetActiveIncidentIDListByShortcutID(incidentValue);
                    var incidentListId = incidentList.LastOrDefault();
                    if (incidentListId != 0)
                    {
                        var incident = CADManager.IncidentQueryEngine.GetIncident(incidentListId);
                        _incidentNumber = incident != null ? incident.IncidentNumber : incidentValue;
                    }
                    else
                    {
                        _incidentNumber = incidentValue;
                    }
                }
                else
                {
                    _incidentNumber = incidentValue;
                }
            }

            _description = Parameters.FirstOrDefault(p => p.Name == ParameterDescription)?.Value;
            _hostName = System.Net.Dns.GetHostName();
            _cadUserId = CADManager.Personnel.PersonnelID;
            _cadChrisId = CADManager.Personnel.Code;
            _cadName = CADManager.Personnel.Name;
            _cadWorkstationId = CADManager.MachineInfoID.ToString();

            // Capture only the first 5 characters of the user string
            if (_cadChrisId is { Length: > 5 })
                _cadChrisId = _cadChrisId.Substring(0, 5);
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while capturing parameters: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }
        return Task.CompletedTask;
    }

    private Task CaptureScreens()
    {
        try
        {
            // Create file path and name
            _timeStamp = DateTime.Now.ToString("yyyyMMdd");
            var fileName = $"CADScreenshot-{_timeStamp}.png";
            var tempPath = Path.GetTempPath();
            var fullPath = Path.Combine(tempPath, fileName);

            // Get the total size of all screens combined (handles multiple monitors)
            var width = SystemInformation.VirtualScreen.Width;
            var height = SystemInformation.VirtualScreen.Height;
            var left = SystemInformation.VirtualScreen.Left;
            var top = SystemInformation.VirtualScreen.Top;

            // Create a bitmap object with the size of the screen
            using var bitmap = new Bitmap(width, height);
            // Use the FromImage method of the Graphics class to create a new Graphics object
            using (var g = Graphics.FromImage(bitmap))
            {
                // Copy the screen contents to the bitmap
                g.CopyFromScreen(left, top, 0, 0, bitmap.Size);
            }

            // Save the screenshot to the specified path
            bitmap.Save(fullPath, ImageFormat.Png);

            _snowScreenshotPath = fullPath;
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while capturing screens: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }
        return Task.CompletedTask;
    }

    private async Task GetSnowUser()
    {
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("x-sn-apikey", _apiKey);

            var employeeId = _cadChrisId;

            var response = await client.GetAsync($"{UserUrl}?sysparm_limit=1&employee_number={employeeId}");

            if (!response.IsSuccessStatusCode)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Request failed with status code: {response.StatusCode}");
            }

            var responseContent = await response.Content.ReadAsStringAsync();
            dynamic serviceNowResponse = JObject.Parse(responseContent);
            _snowUserSysId = serviceNowResponse.SelectToken("result[0].sys_id").ToString();
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while getting the ServiceNow user: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }
    }

    private async Task GetActivityLogs()
    {
        try
        {
            var fileName = $"CADActivityLog-{_timeStamp}.csv";
            var tempPath = Path.GetTempPath();
            var fullPath = Path.Combine(tempPath, fileName);

            var toDate = DateTime.Now;
            var fromDate = toDate.AddMinutes(-30);
            var activityLogData = CADManager.HostedObj.GeneralQueryEngine.GetActivityLogsByPersonnelD(fromDate, toDate, _cadUserId);

            using var writer = new StreamWriter(fullPath);
            // Write the column headers
            await writer.WriteLineAsync("EntryTime,Terminal,IncidentNumber,IncidentID,RadioName,Activity,Comment");

            // Write each row of data
            foreach (var log in activityLogData)
            {
                var comment = log.Comment.Replace("\"", "\"\""); // Escape double quotes
                await writer.WriteLineAsync($"{log.EntryTime},{log.Terminal},{log.IncidentNumber},{log.IncidentID},{log.RadioName},{log.Activity},\"{comment}\"");
            }

            _snowActivityLogPath = fullPath;
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while getting activity logs: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }
    }

    private Task GetLogEntry()
    {
        try
        {
            _snowLogEntriesPath = read_table();
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while getting log entries: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }

        return Task.CompletedTask;
    }

    [DllImport($@"C:\TriTech\VisiCad\NET\logentry.dll")]
    private static extern string read_table();
    
    private async Task CreateSnowTicket()
    {
        try
        {
            using var client = new HttpClient();

            // Set request headers
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("x-sn-apikey", _apiKey);

            var shortDescription = string.IsNullOrEmpty(_incidentNumber) ? $"CAD - System: {_hostName}, Issue: {_description}" : $"CAD - System: {_hostName}, Incident Number: {_incidentNumber}, Issue: {_description}";

            // Create JSON data
            var data = new ServiceNowIncidentData
            {
                caller_id = _snowUserSysId,
                u_department = UserDepartment,
                u_location_string = UserLocation,
                u_street_address = StreetAddress,
                opened_by = _snowUserSysId,
                contact_type = ContactType,
                state = State,
                cmdb_ci = ConfigurationItem,
                assignment_group = AssignmentGroup,
                impact = Impact,
                urgency = Urgency,
                short_description = shortDescription,
                description = $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss},\n" +
                              $"System Hostname: {_hostName},\n" +
                              $"CAD Workstation ID: {_cadWorkstationId},\n" +
                              $"CAD User ID: {_cadUserId},\n" +
                              $"CHRIS ID: {_cadChrisId},\n" +
                              $"User Name: {_cadName},\n" +
                              $"Issue Description: {_description},\n" +
                              $"ServiceNow CAD Integration v2025.3.0"
            };

            var json = JsonConvert.SerializeObject(data);

            var jsonContent = new StringContent(json, Encoding.UTF8, "application/json");

            // Send POST request
            var response = await client.PostAsync(IncidentUrl, jsonContent);
            // Check the response
            if (!response.IsSuccessStatusCode)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Request failed with status code: {response.StatusCode}");
            }

            var responseContent = await response.Content.ReadAsStringAsync();

            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Incident creation response ({response.StatusCode}): {responseContent}");

            dynamic serviceNowResponse = JObject.Parse(responseContent);

            _snowIncidentValue = serviceNowResponse.result.sys_id.ToString();
            
            _emailSubject = "GRIPE: " + data.short_description;
            _emailBody = $"A GRIPE has been submitted by {_cadName} at {_hostName}:\n\n" + 
                         $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss},\n\n" +
                         $"Issue Description: {_description},\n\n" +
                         $"ServiceNow CAD Integration v2025.3.0";
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while creating the ticket: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
            throw;
        }
    }
    
    private async Task UploadFileToServiceNow(string filePath, string tableName, string tableSysId)
    {
        try
        {
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("x-sn-apikey", _apiKey);

            using var content = new MultipartFormDataContent("Upload----" + DateTime.Now.Ticks.ToString("x"));

            // Add table_name and table_sys_id as form fields
            content.Add(new StringContent(tableName), "\"table_name\"");
            content.Add(new StringContent(tableSysId), "\"table_sys_id\"");

            // Read file into byte array
            byte[] fileBytes;
            try
            {
                fileBytes = File.ReadAllBytes(filePath);
            }
            catch (Exception ex)
            {
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Error reading file: {ex.Message}");
                return;
            }

            // Add the file
            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
            fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
            {
                Name = "\"uploadFile\"",
                FileName = "\"" + Path.GetFileName(filePath) + "\""
            };
            content.Add(fileContent);

            // Send the request
            var response = await httpClient.PostAsync(FileUploadUrl, content);

            // Check the response
            var responseContent = await response.Content.ReadAsStringAsync();
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"File upload response ({response.StatusCode}): {responseContent}");
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }
    }
    
    private Task CleanUpGeneratedFiles()
    {
        try
        {
            var filesToDelete = new[] { _snowScreenshotPath, _snowActivityLogPath, _snowLogEntriesPath };

            foreach (var filePath in filesToDelete)
            {
                if (!File.Exists(filePath)) continue;
                File.Delete(filePath);
                CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API",$"Deleted file: {filePath}");
            }
        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"An error occurred while cleaning up files: {ex.Message}");
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Stack trace: {ex.StackTrace}");
        }

        return Task.CompletedTask;
    }

    private Task SendSupervisorEmail()
    {
        var mail = new MailMessage();
        mail.From = new MailAddress("GRIPE@ecscad.local");
        mail.To.Add("ECCSupervisors@cincinnati-oh.gov");
        mail.Subject = _emailSubject;
        mail.Body = _emailBody;
        mail.IsBodyHtml = false;
        
        var smtpClient = new SmtpClient("smtp.rcc.org");
        smtpClient.Port = 25;
        smtpClient.EnableSsl = false;

        try
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API",$"Sending email notification: {_emailSubject}");

            smtpClient.Send(mail);

            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API",$"Email notification sent successfully");

        }
        catch (Exception ex)
        {
            CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API",$"Error sending email notification: {ex.Message}");
        }
        finally
        {
            mail.Dispose(); // Release resources
            smtpClient.Dispose(); // Release resources
        }
        return Task.CompletedTask;
    }

    private void HandleCommandException(Exception ex)
    {
        HandleException(ex);
        m_commandResult.Validations.Add(new CommandValidation(MessageLevel.Error,
            $"Failed to Execute {CommandName} Command: {ex.Message}", string.IsNullOrWhiteSpace(this.UserEnteredText) ? string.Empty : this.UserEnteredText));
        CADManager.GeneralActionEngine.AddActivityLogEntry("ServiceNow API", $"Failed to Execute {CommandName} Command: {ex.Message}");
        m_commandResult.State = CommandState.Failure;
        InvokeCommandComplete(m_commandResult);
    }
    
    // ReSharper disable InconsistentNaming, UnusedMember.Local, UnusedAutoPropertyAccessor.Local, ClassNeverInstantiated.Local
    private class ServiceNowIncidentData
    {
        public string caller_id { get; set; }
        public string u_department { get; set; }
        public string u_location_string { get; set; }
        public string u_street_address { get; set; }
        public string opened_by { get; set; }
        public string contact_type { get; set; }
        public int state { get; set; }
        public string cmdb_ci { get; set; } //Configuration Item
        public string assignment_group { get; set; }
        public int impact { get; set; }
        public int urgency { get; set; }
        public string short_description { get; set; }
        public string description { get; set; }
    }
    // ReSharper restore InconsistentNaming, UnusedMember.Local, UnusedAutoPropertyAccessor.Local, ClassNeverInstantiated.Local
}