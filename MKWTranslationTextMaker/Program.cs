using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Text;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using ArabicUnshapingLib;
using System.Text.Unicode;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace MKWTranslationTextMaker
{
    class Program
    {
        static string language;

        static List<string> commonMIDs;
        static List<string> menuMIDs;
        static List<string> raceMIDs;

        static List<string> appendedCommonMIDs;
        static List<string> appendedMenuMIDs;
        static List<string> appendedRaceMIDs;

        static StringBuilder newCommon;
        static StringBuilder newMenu;
        static StringBuilder newRace;

        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine(System.Diagnostics.Process.GetCurrentProcess().ProcessName + " [language]");
                Exit();
            }

            if (args.Length == 1)
                language = args[0];
            else
                language = string.Join(" ", args);

            string dataPath = Environment.CurrentDirectory + "/data";

            CheckFile(dataPath + "/credentials.txt");
            CheckFile(dataPath + "/Common.txt");
            CheckFile(dataPath + "/Menu.txt");
            CheckFile(dataPath + "/Race.txt");

            string googleClientId = "";
            string googleClientSecret = "";
            string[] scopes = [SheetsService.Scope.Spreadsheets];
            string spreadsheetId = "";

            foreach (string s in File.ReadAllLines(dataPath + "/credentials.txt"))
            {
                if (s.StartsWith("clientId="))
                    googleClientId = s.Split('=')[1];
                else if (s.StartsWith("clientSecret="))
                    googleClientSecret = s.Split('=')[1];
                else if (s.StartsWith("spreadsheetId="))
                    spreadsheetId = s.Split('=')[1];
            }

            UserCredential credential = GoogleAuth.Login(googleClientId, googleClientSecret, scopes);
            GoogleSheetsManager sheetsManager = new GoogleSheetsManager(credential);

            string[] dataCommon = File.ReadAllText(dataPath + "/Common.txt").Split(',');
            string[] dataMenu = File.ReadAllText(dataPath + "/Menu.txt").Split(',');
            string[] dataRace = File.ReadAllText(dataPath + "/Race.txt").Split(',');

            commonMIDs = new List<string>();
            menuMIDs = new List<string>();
            raceMIDs = new List<string>();

            AddMIDs(dataCommon, commonMIDs);
            AddMIDs(dataMenu, menuMIDs);
            AddMIDs(dataRace, raceMIDs);

            appendedCommonMIDs = new List<string>();
            appendedMenuMIDs = new List<string>();
            appendedRaceMIDs = new List<string>();

            newCommon = new StringBuilder("#BMG\r\n");
            newMenu = new StringBuilder("#BMG\r\n");
            newRace = new StringBuilder("#BMG\r\n");

            Spreadsheet spreadsheet = sheetsManager.GetSpreadSheet(spreadsheetId);
            ValueRange valueRange1 = sheetsManager.GetSingleValue(spreadsheet.SpreadsheetId, "'Full Game Translation (Part 1)'!A:ZZZ");
            ValueRange valueRange2 = sheetsManager.GetSingleValue(spreadsheet.SpreadsheetId, "'Full Game Translation (Part 2)'!A:ZZZ");

            AppendTexts(valueRange1);
            AppendTexts(valueRange2);

            AppendEmptyEntries(commonMIDs, appendedCommonMIDs, newCommon);
            AppendEmptyEntries(menuMIDs, appendedMenuMIDs, newMenu);
            AppendEmptyEntries(raceMIDs, appendedRaceMIDs, newRace);

            File.WriteAllText(Environment.CurrentDirectory + "/out/Common.txt", newCommon.ToString());
            File.WriteAllText(Environment.CurrentDirectory + "/out/Menu.txt", newMenu.ToString());
            File.WriteAllText(Environment.CurrentDirectory + "/out/Race.txt", newRace.ToString());
            Exit();
        }

        static void Error(string err)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(err);
            Console.ResetColor();
            Console.WriteLine();
            Exit();
        }

        static void Exit()
        {
            Console.WriteLine("Press Enter to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        static void CheckFile(string path)
        {
            if (!File.Exists(path))
                Error("File not found: " + path);
        }

        static void AddMIDs(string[] array, List<string> list)
        {
            foreach (string s in array)
            {
                if (!list.Contains(s))
                    list.Add(s);
            }
        }

        static void AppendEmptyEntries(List<string> MIDList, List<string> appendedMIDList, StringBuilder builder)
        {
            for (int i = 0; i < MIDList.Count; i++)
            {
                if (!appendedMIDList.Contains(MIDList[i]))
                    builder.AppendLine("   " + MIDList[i] + "\t" + "/");
            }
        }

        static void AppendTexts(ValueRange valueRange)
        {
            int languageColumnIndex = 0;
            for (int i = 0; i < valueRange.Values.Count; i++)
            {
                for (int j = 0; j < valueRange.Values[i].Count; j++)
                {
                    if (i == 0) // Language
                    {
                        if (((string)valueRange.Values[i][j]).ToLower().Equals(language.ToLower()))
                            languageColumnIndex = j;
                        continue;
                    }

                    if (i < 2)
                        continue;

                    if (j != 0) // If not MID column
                        continue;

                    string MID = ((string)valueRange.Values[i][j]).TrimStart('0');
                    string translationText = ((string)valueRange.Values[i][j + languageColumnIndex])
                        .TrimStart(' ').TrimStart('\r').TrimStart('\n').TrimStart(' ')
                        .TrimEnd(' ').TrimEnd('\n').TrimEnd('\r').TrimEnd(' ')
                        .Replace("\r\n", "\n").Replace("\n", "\\n\r\n\t+ ");

                    if (language.ToLower().Equals("arabic"))
                        translationText = ConvertArabic(translationText);

                    string newLine = "   " + MID + "\t" + "= " + translationText;
                    if (commonMIDs.Contains(MID))
                    {
                        newCommon.AppendLine(newLine);
                        appendedCommonMIDs.Add(MID);
                    }
                    if (menuMIDs.Contains(MID))
                    {
                        newMenu.AppendLine(newLine);
                        appendedMenuMIDs.Add(MID);
                    }
                    if (raceMIDs.Contains(MID))
                    {
                        newRace.AppendLine(newLine);
                        appendedRaceMIDs.Add(MID);
                    }
                }
            }
        }

        public static string ReverseString(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        private static string ConvertArabic(string source)
        {
            string arabicWord = string.Empty;
            StringBuilder sbDestination = new StringBuilder();

            bool hasArabicLetters = false;
            foreach (char ch in source)
            {
                if (IsArabic(ch))
                {
                    hasArabicLetters = true;
                    arabicWord += ch;
                }
                else
                {
                    if (arabicWord != string.Empty)
                        sbDestination.Append(ReverseString(Unshaper.DecodeEncodedNonAsciiCharacters(Unshaper.GetUnShapedUnicode(arabicWord))));

                    sbDestination.Append(ch);
                    arabicWord = string.Empty;
                }
            }

            if (!hasArabicLetters)
                return source;

            if (arabicWord != string.Empty)
                sbDestination.Append(ReverseString(Unshaper.DecodeEncodedNonAsciiCharacters(Unshaper.GetUnShapedUnicode(arabicWord))));

            StringBuilder newDest = new StringBuilder();
            string[] lineSplit = sbDestination.ToString().Split("\\n\r\n\t+ ", StringSplitOptions.None);
            foreach (string line in lineSplit)
            {
                List<string> lineWords = line.Split(' ').Reverse().ToList();
                List<string> newLineWords = new List<string>();
                foreach (string word in lineWords)
                {
                    if (string.IsNullOrEmpty(word) || word.Length == 1)
                    {
                        newLineWords.Add(word);
                        continue;
                    }

                    string newWord = word;

                    if (newWord.EndsWith("..."))
                    {
                        newWord = newWord.Remove(newWord.Length - 3);
                        newWord = "..." + newWord;
                    }
                    else
                    {
                        char lastChar = newWord.Last();
                        switch (lastChar)
                        {
                            case '.':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "." + newWord;
                                break;

                            case '?':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "?" + newWord;
                                break;

                            case '!':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "!" + newWord;
                                break;

                            case ':':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = ":" + newWord;
                                break;

                            case '،':
                            case '؟':
                                foreach (char c in word)
                                {
                                    if (c != '،' && c != '؟' && IsArabic(c))
                                        goto afterCheck;
                                }

                                newWord = newWord.Remove(newWord.Length - 1);
                                if (lastChar == '،')
                                    newWord = "،" + newWord;
                                else
                                    newWord = "؟" + newWord;
                                afterCheck:
                                break;

                            /*case '(':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "(" + newWord;
                                break;

                            case ')':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = ")" + newWord;
                                break;

                            case '[':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "[" + newWord;
                                break;

                            case ']':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "]" + newWord;
                                break;

                            case '{':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "{" + newWord;
                                break;

                            case '}':
                                newWord = newWord.Remove(newWord.Length - 1);
                                newWord = "}" + newWord;
                                break;*/
                        }
                    }

                    newWord = Regex.Replace(newWord, @"({)z{(.*,.*)(\\)", @"$3z{$2}");
                    newWord = Regex.Replace(newWord, @"({)u{f(.*)(\\)", @"$3u{f$2}");
                    newWord = newWord.Replace("ls%", "%ls");
                    newLineWords.Add(newWord);
                }
                newDest.Append(string.Join(" ", newLineWords) + "\\n\r\n\t+ ");
            }

            string newStr = newDest.ToString();

            return newStr.Remove(newStr.LastIndexOf("\\n\r\n\t+ "));
        }


        private static bool IsArabic(char character)
        {
            if (character >= 0x600 && character <= 0x6ff)
                return true;

            if (character >= 0x750 && character <= 0x77f)
                return true;

            if (character >= 0xfb50 && character <= 0xfc3f)
                return true;

            if (character >= 0xfe70 && character <= 0xfefc)
                return true;

            return false;
        }
    }
}