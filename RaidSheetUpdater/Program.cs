using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Configuration;
using System.Threading;

namespace RaidSheetUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            ExtractData();
            Thread.Sleep(TimeSpan.FromSeconds(10));
            var champions = LoadChampions();
            var player = ConfigurationSettings.AppSettings["Player"];
            UpdateSheet(champions,player);
        }
        public static void ExtractData()
        {
            System.Diagnostics.Process.Start("Files/Extract.bat");
        }
        public static List<Champions> LoadChampions()
        {
            var champions = new List<Champions>();
            using (StreamReader r = new StreamReader("Files/artifacts.json"))
            {
                string json = r.ReadToEnd();
                dynamic array = JsonConvert.DeserializeObject(json);
                foreach (var item in array.heroes)
                {
                    bool duplicate = false;
                    var championName = (String)item.name;
                    var championGrade = (String)item.grade;
                    champions.ForEach(x => { if (x.name == championName && x.grade == championGrade) duplicate = true; });
                    if (duplicate)
                    {
                        //Console.WriteLine("Duplicate entry found: " + championName + " " + championGrade);
                    }
                    else
                    {
                        champions.Add(new Champions(championName, championGrade));
                    }
                }
            }
            return champions;
        }

        public static void UpdateSheet(List<Champions> champions, string player)
        {
            ServiceAccountCredential credential;
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string serviceAccountEmail = "raidimportjson@import-raid.iam.gserviceaccount.com";
            string jsonfile = "Files/import-raid-382cfa628211.json";
            string ApplicationName = "RaidSL";
            using (Stream stream = new FileStream(@jsonfile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                credential = (ServiceAccountCredential)
                    GoogleCredential.FromStream(stream).UnderlyingCredential;

                var initializer = new ServiceAccountCredential.Initializer(credential.Id)
                {
                    User = serviceAccountEmail,
                    Key = credential.Key,
                    Scopes = Scopes
                };
                credential = new ServiceAccountCredential(initializer);

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                var startRow = 3;
                var totalHeros = champions.Count();
                var endRow = startRow + totalHeros;
                string spreadsheetId = "1IRGd_c6bPk94PwSeT3fNn0dN-Z0GKty524ACmu_TU-U";
                string championColumn = "";
                string starColumn = "";
                if (player.ToLower() == "paul")
                {
                    championColumn = "A";
                    starColumn = "B";
                }
                if (player.ToLower() == "ruan")
                {
                    championColumn = "F";
                    starColumn = "G";
                }
                if (player.ToLower() == "nico")
                {
                    championColumn = "K";
                    starColumn = "L";
                }
                if (player.ToLower() == "patric")
                {
                    championColumn = "P";
                    starColumn = "Q";
                }

                //Clear rows
                string clearChampions = "RaidImport!"+ championColumn + "3:"+ championColumn + "1002";
                string clearStars = "RaidImport!"+ starColumn + "3:"+ starColumn + "1002";
                ValueRange rangeChampionsClear = new ValueRange();
                ValueRange rangeStarsClear = new ValueRange();
                rangeChampionsClear.MajorDimension = "COLUMNS";
                rangeStarsClear.MajorDimension = "COLUMNS";
                var championListClear = new List<object>();
                var starListClear = new List<object>();
                for (var i = 0; i < 1000;i++)
                {
                    championListClear.Add("");
                    starListClear.Add("");
                }

                rangeChampionsClear.Values = new List<IList<object>> { championListClear };
                rangeStarsClear.Values = new List<IList<object>> { starListClear };

                SpreadsheetsResource.ValuesResource.UpdateRequest updateClear = service.Spreadsheets.Values.Update(rangeChampionsClear, spreadsheetId, clearChampions);
                SpreadsheetsResource.ValuesResource.UpdateRequest updateClear2 = service.Spreadsheets.Values.Update(rangeStarsClear, spreadsheetId, clearStars);
                updateClear.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateClear2.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse resultClear = updateClear.Execute();
                UpdateValuesResponse resultClear2 = updateClear2.Execute();

                string updateChampions = "RaidImport!"+ championColumn + startRow.ToString() + ":"+ championColumn + endRow.ToString();
                string updateStars = "RaidImport!"+ starColumn + startRow.ToString() + ":"+ starColumn + endRow.ToString();
                ValueRange valueRange = new ValueRange();
                ValueRange valueRange2 = new ValueRange();
                valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS
                valueRange2.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

                var championList = new List<object>();
                var starList = new List<object>();
                champions.ForEach(x => { championList.Add(x.name); starList.Add(x.grade); });
                valueRange.Values = new List<IList<object>> { championList };
                valueRange2.Values = new List<IList<object>> { starList };

                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, updateChampions);
                SpreadsheetsResource.ValuesResource.UpdateRequest update2 = service.Spreadsheets.Values.Update(valueRange2, spreadsheetId, updateStars);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                update2.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result = update.Execute();
                UpdateValuesResponse result2 = update2.Execute();

                Console.WriteLine("done!");
            }
        }

    }

    class Champions
    {
        public string name { get; set; }
        public string grade { get; set; }
        public Champions(string name, string grade)
        {
            this.name = name;
            this.grade = grade;
        }
    }
}
