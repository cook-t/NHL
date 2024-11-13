using Aspose.Cells;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using Nhl.Api;
using Nhl.Api.Models.Team;
using Nhl.Api.Models.Game;
using System.Net.Mail;
using System.Net;
using System.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using Nhl.Api.Models.Schedule;

using HttpClient client = new();
client.DefaultRequestHeaders.Accept.Clear();
client.DefaultRequestHeaders.Add("key", "265c6c78a53941dd9e35ce94b29f1049");

SqlConnection conn = new SqlConnection("Data Source=localhost\\SQLEXPRESS;Initial Catalog=NHL;Integrated Security=SSPI;");

var nhl = new NhlApi();
List<Nhl.Api.Models.Team.Teams> teams = new();

var nhlGame = new NhlGameApi();

bool runForTomorrow = false;

using (var webClient = new WebClient())
{
    Console.WriteLine("Downloading all_teams.csv...");
    webClient.DownloadFile("https://moneypuck.com/moneypuck/playerData/careers/gameByGame/all_teams.csv", @"C:\Users\cooktyl\Documents\NHL\all_teams.csv");
    Console.WriteLine("Download complete.");
}

Console.WriteLine("Getting today's games...");
DateOnly slateDate = (!runForTomorrow) ? DateOnly.FromDateTime(DateTime.Now) : DateOnly.FromDateTime(DateTime.Now.AddDays(1));
var gamesWeek = await nhl.GetLeagueGameWeekScheduleByDateAsync(slateDate);
List<Nhl.Api.Models.Schedule.Game> todaysGames = gamesWeek.GameWeek[0].Games;

//GetAllTeamsList();

SortedDictionary<int, string> GetAllTeamsList()
{
    List<Nhl.Api.Models.Schedule.Game> tomorrowsGames = gamesWeek.GameWeek[1].Games;
    List<Nhl.Api.Models.Schedule.Game> twodaysGames = gamesWeek.GameWeek[2].Games;
    List<Nhl.Api.Models.Schedule.Game> threedaysGames = gamesWeek.GameWeek[3].Games;
    List<Nhl.Api.Models.Schedule.Game> fourdaysGames = gamesWeek.GameWeek[4].Games;
    List<Nhl.Api.Models.Schedule.Game> fivedaysGames = gamesWeek.GameWeek[5].Games;

    var allGames = new List<Nhl.Api.Models.Schedule.Game>();
    allGames = todaysGames.Concat(tomorrowsGames).Concat(twodaysGames).Concat(threedaysGames).Concat(fourdaysGames).Concat(fivedaysGames).ToList();

    SortedDictionary<int, string> teamDict = new();

    foreach (Nhl.Api.Models.Schedule.Game game in allGames)
    {
        string homeTeamId;
        string awayTeamId;

        if (teamDict.TryGetValue(game.HomeTeam.Id, out homeTeamId))
        {
            continue;
        }
        else
        {
            teamDict.Add(game.HomeTeam.Id, game.HomeTeam.Abbrev);
        }

        if (teamDict.TryGetValue(game.AwayTeam.Id, out awayTeamId))
        {
            continue;
        }
        else
        {
            teamDict.Add(game.AwayTeam.Id, game.AwayTeam.Abbrev);
        }
    }

    return teamDict;
}


Console.WriteLine("Games retrieved.");

DataTable dtPriorGameAverages = new();
DataTable csvData = new();

Console.WriteLine("Starting data processing...");
csvData = GetDataTabletFromCSVFile(@"C:\Users\cooktyl\Documents\NHL\all_teams.csv");
UpdateGameDateFormat();
TruncateAllTeamsTable();
InsertDataIntoSQLServerUsingSQLBulkCopy(csvData);
Console.WriteLine("Data processing complete.");

Console.WriteLine("Getting prior matchup data...");
GetPriorMatchupData();
Console.WriteLine("Prior matchup data retrieved.");

Console.WriteLine("Building Excel file...");
BuildExcelFile();
Console.WriteLine("Excel file built.");

void UpdateGameDateFormat()
{
    foreach (DataRow dr in csvData.Rows)
    {
        string origValue = dr["gameDate"].ToString();
        string newValue = origValue.Substring(4, 2) + "/" + origValue.Substring(6, 2) + "/" + origValue.Substring(0, 4);
        dr["gameDate"] = newValue;
    }
}

DataTable GetDataTabletFromCSVFile(string csv_file_path)
{
    DataTable csvData = new DataTable();
    try
    {
        using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
        {
            csvReader.SetDelimiters(new string[] { "," });
            csvReader.HasFieldsEnclosedInQuotes = true;
            string[] colFields = csvReader.ReadFields();
            foreach (string column in colFields)
            {
                DataColumn datecolumn = new DataColumn(column);
                datecolumn.AllowDBNull = true;
                csvData.Columns.Add(datecolumn);
            }
            while (!csvReader.EndOfData)
            {
                string[] fieldData = csvReader.ReadFields();
                //Making empty value as null
                for (int i = 0; i < fieldData.Length; i++)
                {
                    if (fieldData[i] == "")
                    {
                        fieldData[i] = null;
                    }
                }
                csvData.Rows.Add(fieldData);
            }
        }
    }
    catch (Exception ex)
    {
        return null;
    }
    return csvData;
}

void TruncateAllTeamsTable()
{
    using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE all_teams", conn) { CommandType = CommandType.Text })
    {
        conn.Open();
        cmd.ExecuteNonQuery();
        conn.Close();
    }
}

void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData)
{
    try
    {
        using (SqlBulkCopy s = new SqlBulkCopy(conn))
        {
            conn.Open();
            s.DestinationTableName = "all_teams";
            foreach (var column in csvFileData.Columns)
                s.ColumnMappings.Add(column.ToString(), column.ToString());
            s.WriteToServer(csvFileData);
            conn.Close();
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
}

async void GetPriorMatchupData()
{
    int gameCount = 10;

    using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE OutputData", conn) { CommandType = CommandType.Text })
    {
        conn.Open();
        cmd.ExecuteNonQuery();
    }

    foreach (Nhl.Api.Models.Schedule.Game game in todaysGames)
    {
        using (SqlCommand cmd = new SqlCommand("GetTeamPriorData", conn) { CommandType = CommandType.StoredProcedure })
        {
            cmd.Parameters.Add("@TEAMID", SqlDbType.Int);
            cmd.Parameters.Add("@HOMEAWAY", SqlDbType.Text);
            cmd.Parameters.Add("@GAMECOUNT", SqlDbType.Int);

            cmd.Parameters["@TEAMID"].Value = game.AwayTeam.Id;
            cmd.Parameters["@HOMEAWAY"].Value = "AWAY";
            cmd.Parameters["@GAMECOUNT"].Value = gameCount;
            cmd.ExecuteNonQuery();

            cmd.Parameters["@TEAMID"].Value = game.HomeTeam.Id;
            cmd.Parameters["@HOMEAWAY"].Value = "HOME";
            cmd.Parameters["@GAMECOUNT"].Value = gameCount;
            cmd.ExecuteNonQuery();
        }
    }

    using (SqlCommand cmd = new SqlCommand("SELECT * FROM OutputData", conn) { CommandType = CommandType.Text })
    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
    {
        da.Fill(dtPriorGameAverages);
    }

    conn.Close();
}

void BuildExcelFile()
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    string date = slateDate.Month.ToString() + "_" + slateDate.Day.ToString() + "_" + slateDate.Year.ToString();
    string outputFileDir = @"C:\Users\cooktyl\Documents\NHL\NHL_" + date + ".xlsx";
    string templateFileDir = @"C:\Users\cooktyl\Documents\NHL\NHLTemplate.xlsx";

    FileInfo newFile = new FileInfo(outputFileDir);
    FileInfo tempFile = new FileInfo(templateFileDir);

    FileInfo chkExistingFile = new FileInfo(outputFileDir);
    if (chkExistingFile.Exists)
    {
        chkExistingFile.Delete();
    }

    try
    {
        using (ExcelPackage pck = new ExcelPackage(newFile, tempFile))
        {
            int priorMatchupSheetCellNum = 1;
            dtPriorGameAverages.Columns.RemoveAt(0);
            ExcelWorksheet ws2 = pck.Workbook.Worksheets["Data"];
            ExcelWorksheet ws3 = pck.Workbook.Worksheets["PriorMatchupAvgs"];
            ws2.Cells["A1"].LoadFromDataTable(dtPriorGameAverages, true);

            //PriorMatchups
            foreach (Nhl.Api.Models.Schedule.Game game in todaysGames)
            {

                DataTable dtPriorMatchupAverages = new();
                using (SqlCommand cmd = new SqlCommand("TeamsMatchupHistory", conn) { CommandType = CommandType.StoredProcedure })
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.Add("@TEAM1ID", SqlDbType.Int);
                    cmd.Parameters.Add("@TEAM2ID", SqlDbType.Int);

                    cmd.Parameters["@TEAM1ID"].Value = game.HomeTeam.Id;
                    cmd.Parameters["@TEAM2ID"].Value = game.AwayTeam.Id;
                    da.Fill(dtPriorMatchupAverages);
                    ws3.Cells["A" + priorMatchupSheetCellNum.ToString()].LoadFromDataTable(dtPriorMatchupAverages, true);
                }

                priorMatchupSheetCellNum = priorMatchupSheetCellNum + 2;
            }

            pck.Save();
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
}


string testpoint = "";

//var smtpClient = new SmtpClient("smtp.gmail.com")
//{
//    Port = 587,
//    Credentials = new NetworkCredential("tylercooked@gmail.com", "SteelPanther_05"),
//    EnableSsl = true,
//};

//using (MailMessage message = new MailMessage())
//{
//    message.To.Add("tylercooked@gmail.com");
//    message.From = new MailAddress("DoNotReply@NHLData.com");
//    message.Subject = "Today's NHL Data";
//    Attachment workbook = new Attachment(@"C:\Users\cooktyl\Documents\NHL\NHL_" + date + ".xlsx");
//    message.Attachments.Add(workbook);

//    smtpClient.Send(message);
//}