using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ApiTools.Json;
using ApiTools.Net;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;

namespace HelloWorld
{
    public partial class SprintRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Btn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var worksheet = Globals.ThisAddIn.Application.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Team");

                if (worksheet == null)
                {
                    MessageBox.Show("Error! Could not find 'Team' worksheet. Is this a sprint planning workbook?");
                }
                else
                {
                    var cell = (Excel.Range)worksheet.Cells[2][4];
                    var sprintNumber = cell.Value.ToString();

                    var fetchUrl = @"http://ca01a2626:4040/tasks/sprints/" + sprintNumber + @"/hours";
                    var client = new RestClient();
                    var response = client.Fetch(fetchUrl);

                    if (response.Code == 200)
                    {
                        var membersJson = new JsonArray(response.Body);
                        var teamMembers = GetTeamMembers(worksheet);
                        UpdatePlannedHours(teamMembers, membersJson);
                        UpdateWorksheet(teamMembers, worksheet);
                    }
                    else
                    {
                        MessageBox.Show("Failed to contact JIRA at " + fetchUrl);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("INTERNAL ERROR: " + ex.Message);
            }
        }

        private void UpdateWorksheet(List<TeamMember> members, Excel.Worksheet worksheet)
        {
            foreach (var member in members)
            {
                member.UpdatePlannedHours(worksheet);
            }
        }

        private void UpdatePlannedHours(List<TeamMember> members, JsonArray memersJson)
        {
            foreach (JsonKeyValuePairs kvp in memersJson)
            {
                var id = (JsonKeyValuePairs)kvp["_id"];
                var username = (string)id["assignee"];
                var plannedHoursString = kvp["remaining"].ToString();
                var plannedHours = double.Parse(plannedHoursString);

                var teamMember = members.FirstOrDefault(x => x.Username == username);
                if (teamMember != null)
                {
                    teamMember.PlannedHours = plannedHours;
                }
            }
        }

        private List<TeamMember> GetTeamMembers(Excel.Worksheet worksheet)
        {
            const int FirstColumn = 3;
            const int FirstRow = 19;
            int currentRow = FirstRow;
            var users = new List<TeamMember>();

            while(true)
            {
                var username = (string)(worksheet.Cells[FirstColumn][currentRow] as Excel.Range).Value;

                if (!string.IsNullOrEmpty(username))
                {
                    users.Add(new TeamMember(currentRow, FirstColumn, username));
                    currentRow++;
                }
                else
                {
                    break;
                }

                System.Diagnostics.Debug.Write(username);
            }

            return users;
        }
    }
}
