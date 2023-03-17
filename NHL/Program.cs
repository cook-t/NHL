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

//using HttpClient client = new();
//client.DefaultRequestHeaders.Accept.Clear();
//client.DefaultRequestHeaders.Add("key", "265c6c78a53941dd9e35ce94b29f1049");

//await ProcessRepositoriesAsync(client);

//Aspose.Cells.License license = new Aspose.Cells.License();
//license.SetLicense(Assembly.GetExecutingAssembly().GetManifestResourceStream("NHL.Other.Aspose.Total.lic"));

//static async Task ProcessRepositoriesAsync(HttpClient client)
//{
//    var json = await client.GetStringAsync(
//         "https://api.sportsdata.io/v3/nhl/scores/json/GamesByDate/2023-FEB-17?key=265c6c78a53941dd9e35ce94b29f1049");

//    DataTable dt = (DataTable)JsonConvert.DeserializeObject(json, (typeof(DataTable)));
//    dt.DefaultView.Sort = "DateTime desc";
//    DataTable dtSorted = dt.DefaultView.ToTable();
//    string fileDir = @"C:\Users\cooktyl\Documents\NHL\NHL.xlsx";

//    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//    using (ExcelPackage pck = new ExcelPackage(@"C:\Users\cooktyl\Documents\NHL\NHL.xlsx"))
//    {
//        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Data");
//        ws.Cells["A1"].LoadFromDataTable(dtSorted, true);
//        pck.Save();
//    }
//}

SqlConnection conn = new SqlConnection("Data Source=localhost\\SQLEXPRESS;Initial Catalog=NHL;Integrated Security=SSPI;");

var nhl = new NhlApi();
List<Nhl.Api.Models.Team.Team> teams = new();

var nhlGame = new NhlGameApi();

GameSchedule todaysGames = new();
List<int> awayTeams = new();
List<int> homeTeams = new();
DataTable dt = new();
dt.Columns.Add("TeamID", typeof(Int32));
dt.Columns.Add("TeamName", typeof(String));
dt.Columns.Add("AvgGoals", typeof(Double));
dt.Columns.Add("AvgShots", typeof(Double));
dt.Columns.Add("AvgGoalsAgainst", typeof(Double));
dt.Columns.Add("AvgShotsAgainst", typeof(Double));
dt.Columns.Add("AvgTotalGoals", typeof(Double));

DataTable dtTodaysTeamAvgGoals = new();
dtTodaysTeamAvgGoals.Columns.Add("AwayTeamID", typeof(Int32));
dtTodaysTeamAvgGoals.Columns.Add("AwayTeamName", typeof(String));
dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgGoals", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgGoalsAgainst", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgShots", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgShotsAgainst", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgTotalGoals", typeof(Double));

//dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgGoals", typeof(Double));
//dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgTotalGoals", typeof(Double));
//dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgShots", typeof(Double));
//dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgShotsAgainst", typeof(Double));
//dtTodaysTeamAvgGoals.Columns.Add("L10AwayAvgGoalsAgainst", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("HomeTeamID", typeof(Int32));
dtTodaysTeamAvgGoals.Columns.Add("HomeTeamName", typeof(String));
dtTodaysTeamAvgGoals.Columns.Add("L10HomeAvgGoals", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10HomeAvgGoalsAgainst", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10HomeAvgShots", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10HomeAvgShotsAgainst", typeof(Double));
dtTodaysTeamAvgGoals.Columns.Add("L10HomeAvgTotalGoals", typeof(Double));



teams = await nhl.GetActiveTeamsAsync();

//using (SqlCommand cmd = new SqlCommand("SET IDENTITY_INSERT Teams ON INSERT INTO Teams ([TeamID], [Name], [Abbreviation], [City], [TeamName])  VALUES (@TeamID, @Name, @Abbreviation, @City, @TeamName)", conn))
//{
//    conn.Open();
//    cmd.Parameters.Add("@TeamID", SqlDbType.Int);
//    cmd.Parameters.Add("@Name", SqlDbType.Text);
//    cmd.Parameters.Add("@Abbreviation", SqlDbType.Text);
//    cmd.Parameters.Add("@City", SqlDbType.Text);
//    cmd.Parameters.Add("@TeamName", SqlDbType.Text);

//    foreach (Team team in teams)
//    {
//        cmd.Parameters["@TeamID"].Value = team.Id;
//        cmd.Parameters["@Name"].Value = team.Name;
//        cmd.Parameters["@Abbreviation"].Value = team.Abbreviation;
//        cmd.Parameters["@City"].Value = team.LocationName;
//        cmd.Parameters["@TeamName"].Value = team.TeamName;
//        cmd.ExecuteNonQuery();
//    }
//}


todaysGames = await nhlGame.GetGameScheduleByDateAsync(DateTime.Now);
SplitTodaysTeams();
await GetTeamAverages(10, "home");
await GetTeamAverages(10, "away");
//await GetTeamAverages(10, null);
await InsertAvgGoalValues();


async void SplitTodaysTeams()
{
    foreach (Nhl.Api.Models.Game.Game game in todaysGames.Dates[0].Games)
    {
        awayTeams.Add(game.Teams.AwayTeam.Team.Id);
        homeTeams.Add(game.Teams.HomeTeam.Team.Id);

        DataRow dr = dtTodaysTeamAvgGoals.NewRow();
        dr["AwayTeamID"] = game.Teams.AwayTeam.Team.Id;
        dr["AwayTeamName"] = game.Teams.AwayTeam.Team.Name;
        dr["HomeTeamID"] = game.Teams.HomeTeam.Team.Id;
        dr["HomeTeamName"] = game.Teams.HomeTeam.Team.Name;

        dtTodaysTeamAvgGoals.Rows.Add(dr);
    }
}





async Task<List<Game>> GetTeamsPreviousGames(int teamID, string homeaway, int amountOfGames)
{
    Nhl.Api.Models.Game.GameSchedule priorGames = await nhlGame.GetGameScheduleForTeamByDateAsync(teamID, new DateTime(2022, 10, 07), DateTime.Now.AddDays(-1));
    List<GameDate> gameDateBatch = new List<GameDate>();
    List<Game> games = new();
    
    if (homeaway == null)
    {
        gameDateBatch = priorGames.Dates;
    }
    if (homeaway == "away")
    {
        foreach (GameDate gameDate in priorGames.Dates)
        {
            if (gameDate.Games[0].Teams.AwayTeam.Team.Id == teamID)
            {
                gameDateBatch.Add(gameDate);
            }
        }
    }
    if (homeaway == "home")
    {
        foreach (GameDate gameDate in priorGames.Dates)
        {
            if (gameDate.Games[0].Teams.HomeTeam.Team.Id == teamID)
            {
                gameDateBatch.Add(gameDate);
            }
        }
    }

    int cntGames = gameDateBatch.Count;

    for (int i = cntGames - amountOfGames; i < cntGames; i++)
    {
        games.Add(gameDateBatch[i].Games[0]);
    }

    return games;
}

async Task GetTeamAverages(int numberOfGames, string homeaway)
{
    if (homeaway == "home")
    {
        foreach (int teamID in homeTeams)
        {
            List<Game> prevGames = new();
            prevGames = await GetTeamsPreviousGames(teamID, homeaway, numberOfGames);
            List<int> goals = new();
            List<int> goalsAgainst = new();
            List<int> totalGoals = new();
            List<int> shots = new();
            List<int> shotsAgainst = new();

            foreach (Game game in prevGames)
            {
                Nhl.Api.Models.Game.Boxscore boxScore = await nhlGame.GetBoxScoreByIdAsync(game.GamePk);
                goals.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals);
                shots.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Shots);
                totalGoals.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals + boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                goalsAgainst.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                shotsAgainst.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Shots);
            }

            double avgGoals = goals.Average();
            double avgTotalGoals = totalGoals.Average();
            double avgGoalsAgainst = goalsAgainst.Average();
            double avgShotsAgainst = shotsAgainst.Average();
            double avgShots = shots.Average();
            DataRow dr = dt.NewRow();
            dr["TeamID"] = teamID;
            dr["TeamName"] = teams.Where(t => t.Id == teamID).FirstOrDefault().Name;
            dr["AvgGoals"] = avgGoals;
            dr["AvgTotalGoals"] = avgTotalGoals;
            dr["AvgGoalsAgainst"] = avgGoalsAgainst;
            dr["AvgShotsAgainst"] = avgShotsAgainst;
            dr["AvgShots"] = avgShots;
            dt.Rows.Add(dr);
        }
    }
    else if (homeaway == "away")
    {
        foreach (int teamID in awayTeams)
        {
            List<Game> prevGames = new();
            prevGames = await GetTeamsPreviousGames(teamID, homeaway, numberOfGames);
            List<int> goals = new();
            List<int> goalsAgainst = new();
            List<int> totalGoals = new();
            List<int> shots = new();
            List<int> shotsAgainst = new();

            foreach (Game game in prevGames)
            {
                Nhl.Api.Models.Game.Boxscore boxScore = await nhlGame.GetBoxScoreByIdAsync(game.GamePk);
                goals.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                shots.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Shots);
                totalGoals.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals + boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                goalsAgainst.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals);
                shotsAgainst.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Shots);
            }

            double avgGoals = goals.Average();
            double avgTotalGoals = totalGoals.Average();
            double avgGoalsAgainst = goalsAgainst.Average();
            double avgShotsAgainst = shotsAgainst.Average();
            double avgShots = shots.Average();
            DataRow dr = dt.NewRow();
            dr["TeamID"] = teamID;
            dr["TeamName"] = teams.Where(t => t.Id == teamID).FirstOrDefault().Name;
            dr["AvgGoals"] = avgGoals;
            dr["AvgTotalGoals"] = avgTotalGoals;
            dr["AvgGoalsAgainst"] = avgGoalsAgainst;
            dr["AvgShotsAgainst"] = avgShotsAgainst;
            dr["AvgShots"] = avgShots;
            dt.Rows.Add(dr);
        }
    }
    else
    {
        List<int> allTeams = new();
        foreach (int teamID in awayTeams)
        {
            allTeams.Add(teamID);
        }
        foreach (int teamID in homeTeams)
        {
            allTeams.Add(teamID);
        }
        foreach (int teamID in allTeams)
        {
            List<Game> prevGames = new();
            prevGames = await GetTeamsPreviousGames(teamID, null, numberOfGames);
            List<int> goals = new();
            List<int> goalsAgainst = new();
            List<int> totalGoals = new();
            List<int> shots = new();
            List<int> shotsAgainst = new();

            foreach (Game game in prevGames)
            {
                Nhl.Api.Models.Game.Boxscore boxScore = await nhlGame.GetBoxScoreByIdAsync(game.GamePk);
                if (boxScore.Teams.Away.TeamInformation.Id == teamID)
                {
                    goals.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                    shots.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Shots);
                    goalsAgainst.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals);
                    shotsAgainst.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Shots);
                }
                if (boxScore.Teams.Home.TeamInformation.Id == teamID)
                {
                    goals.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals);
                    shots.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Shots);
                    goalsAgainst.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
                    shotsAgainst.Add(boxScore.Teams.Away.TeamStats.TeamSkaterStats.Shots);
                }
                totalGoals.Add(boxScore.Teams.Home.TeamStats.TeamSkaterStats.Goals + boxScore.Teams.Away.TeamStats.TeamSkaterStats.Goals);
            }

            double avgGoals = goals.Average();
            double avgTotalGoals = totalGoals.Average();
            double avgGoalsAgainst = goalsAgainst.Average();
            double avgShotsAgainst = shotsAgainst.Average();
            double avgShots = shots.Average();
            DataRow dr = dt.NewRow();
            dr["TeamID"] = teamID;
            dr["TeamName"] = teams.Where(t => t.Id == teamID).FirstOrDefault().Name;
            dr["AvgGoals"] = avgGoals;
            dr["AvgTotalGoals"] = avgTotalGoals;
            dr["AvgGoalsAgainst"] = avgGoalsAgainst;
            dr["AvgShotsAgainst"] = avgShotsAgainst;
            dr["AvgShots"] = avgShots;
            dt.Rows.Add(dr);
        }
    }
}

async Task InsertAvgGoalValues()
{
    foreach (DataRow dr in dtTodaysTeamAvgGoals.Rows)
    {
        foreach (DataRow dr2 in dt.Rows)
        {
            if (Convert.ToInt32(dr2["TeamID"]) == Convert.ToInt32(dr["AwayTeamID"]))
            {
                dr["L10AwayAvgGoals"] = dr2["AvgGoals"];
                dr["L10AwayAvgTotalGoals"] = dr2["AvgTotalGoals"];
                dr["L10AwayAvgShots"] = dr2["AvgShots"];
                dr["L10AwayAvgGoalsAgainst"] = dr2["AvgGoalsAgainst"];
                dr["L10AwayAvgShotsAgainst"] = dr2["AvgShotsAgainst"];
            }
            else if (Convert.ToInt32(dr2["TeamID"]) == Convert.ToInt32(dr["HomeTeamID"]))
            {
                dr["L10HomeAvgGoals"] = dr2["AvgGoals"];
                dr["L10HomeAvgTotalGoals"] = dr2["AvgTotalGoals"];
                dr["L10HomeAvgShots"] = dr2["AvgShots"];
                dr["L10HomeAvgGoalsAgainst"] = dr2["AvgGoalsAgainst"];
                dr["L10HomeAvgShotsAgainst"] = dr2["AvgShotsAgainst"];
            }
        }
    }
}

//string fileDir = @"C:\Users\cooktyl\Documents\NHL\NHL.xlsx";

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string date = DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Year.ToString();

using (ExcelPackage pck = new ExcelPackage(@"C:\Users\cooktyl\Documents\NHL\NHL_" + date + ".xlsx"))
{
    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Data");
    ws.Cells["A1"].LoadFromDataTable(dtTodaysTeamAvgGoals, true);
    pck.Save();
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