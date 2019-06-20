using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using urlReader.Classes;

namespace urlReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //"http://([\\w+?\\.\\w+])+([a-zA-Z0-9\\~\\!\\@\\#\\$\\%\\^\\&amp;\\*\\(\\)_\\-\\=\\+\\\\\\/\\?\\.\\:\\;\\'\\,]*)?", RegexOptions.IgnoreCase
            //@"http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
            var linkParser = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);


            string redhatFilePath = Constants.Constants.redhatFile;
            string suseFilePath = Constants.Constants.suseFile;

            var dictRedHat = new Dictionary<string, string>();
            var dictSuse = new Dictionary<string, string>();




            FileInfo redhatFile = new FileInfo(redhatFilePath);
            FileInfo suseFile = new FileInfo(suseFilePath);
            string html = string.Empty;

            FileInfo[] files = new FileInfo[] { redhatFile, suseFile };


            dictRedHat = CreateAndFillTHeValuesForRedHat(redhatFile);
            dictSuse = CreateAndFillTheValuesForSuse(suseFile);

            FillRedHatFile(redhatFile, dictRedHat);
            //FillSuseFile(suseFile, dictSuse);



        }

        private static void FillSuseFile(FileInfo suseFile, Dictionary<string, string> dictSuse)
        {
            Dictionary<string, List<string>> dictWithPackages = new Dictionary<string, List<string>>();
            List <NotParsedSuse> listNotParsedSuse = new List<NotParsedSuse>();
            Dictionary<string, List<Package>> Packages = new Dictionary<string, List<Package>>();

            List<Suse> dataFile = new List<Suse>();

            using (var package = new ExcelPackage(suseFile))
            {
                ExcelWorksheet rawData = package.Workbook.Worksheets[1];

                int rows = rawData.Dimension.Rows;

                //change i to 2 after the debug is over
                for (int i = 2; i <= rows; i++)
                {
                    Console.WriteLine(i);
                    string bulletinId = Convert.ToString(rawData.Cells[i, 1].Value);
                    string affectedPackages = dictSuse[bulletinId];
                    string bulletinTitle = Convert.ToString(rawData.Cells[i, 2].Value);

                    Packages.Add(bulletinId, new List<Package>());

                    // Remove tab spaces
                    affectedPackages = affectedPackages.Replace("\t", " ");


                    // Remove multiple white spaces from HTML
                    //affectedPackages = Regex.Replace(affectedPackages, "\\s+", " ");


                    affectedPackages = Regex.Replace(affectedPackages, "<[^>]*>", "");

                    string[] parsedValues = affectedPackages.Split(new string[] { "          ", "            ", "  ", "\t\t\t", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    string lastAddedKey = "";

                    foreach (var item in parsedValues.Skip(1))
                    {
                        

                        if (item.StartsWith(" - SUSE"))
                        {
                            
                            dictWithPackages.Add(item, new List<string>());
                            lastAddedKey = item;
                           
                        }
                        else
                        {
                           
                            dictWithPackages[lastAddedKey].Add(item);
                        }
                    }



                    var suse = new Suse()
                    {
                        name = bulletinId,
                        Title = bulletinTitle,
                        Packages = new List<Package>()


                    };

                    foreach (var element in dictWithPackages)
                    {
                        string []arrayWithTitleAndVersions = element.Key.Split(new char[] { '(' }, StringSplitOptions.RemoveEmptyEntries);
                        string title = arrayWithTitleAndVersions[0].Replace('-', ' ').TrimStart(' ').TrimEnd(' ');
                        string versionsRaw = arrayWithTitleAndVersions[1];
                        string versionsUpdated = versionsRaw.Remove(versionsRaw.Length - 2, 2);
                       
                        var versions = versionsUpdated.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                        var newPackage = new Package()
                        {
                            Name = title,
                            Versions = new List<string>(),
                            Components = new List<string>()
                        };
                        newPackage.Versions.AddRange(versions);
                        newPackage.Components.AddRange(element.Value);
                        Packages[bulletinId].Add(newPackage);
                        suse.Packages.Add(newPackage);

                        dataFile.Add(suse);

                        


                    }
                    dictWithPackages.Clear();

                    string outputValues = ComposeSuseOutputForExcel(parsedValues);

                    rawData.Cells[i, 9].Value = outputValues;
                   

                }

                package.Save();
            

                ExcelWorksheet dataSheet = package.Workbook.Worksheets.Add("SuseData");

                int row = 2;
                int col = 1;

                foreach (var advisory in Packages)
                {
                    string advisoryName = advisory.Key;

                    foreach (var packageItem in advisory.Value)
                    {
                        string versionsInfo = string.Join(" ", packageItem.Versions);

                        foreach (var component in packageItem.Components)
                        {
                            string mergedInfo = $"{advisoryName}|{packageItem.Name}|{versionsInfo}|{component}";
                            dataSheet.Cells[row++, col].Value = mergedInfo;
                        }
 
                        
                    }
                }

                //int currentRow = 1;
                //int currentCol = 1;

                //foreach (var item in dataFile)
                //{
                //    string name = item.name;
                //    string title = item.Title;

                //    foreach (var pack in item.Packages)
                //    {
                //        string versions = string.Join(" ", pack.Versions);

                //        foreach (var component in pack.Components)
                //        {
                //            dataSheet.Cells[currentRow, currentCol++].Value = name;
                //            dataSheet.Cells[currentRow, currentCol++].Value = pack.Name;
                //            dataSheet.Cells[currentRow, currentCol++].Value = versions;
                //            dataSheet.Cells[currentRow, currentCol++].Value = component;
                //            string mergedInfo = $"{name}|{pack.Name}|{versions}|{component}";
                //            dataSheet.Cells[currentRow, currentCol].Value = mergedInfo;
                //            currentRow++;
                //            currentCol = 1;

                //        }
                //    }
                //}

                package.Save();
                package.Dispose();
            }
        }

        private static string ComposeSuseOutputForExcel(string[] parsedValues)
        {
            return string.Join("|", parsedValues.Take(parsedValues.Length-1).Skip(1));
        }

        private static void FillRedHatFile(FileInfo redhatFile, Dictionary<string, string> dictRedHat)
        {
            var dict = new Dictionary<string, Dictionary<string,List<string>>>();
            int rowIndex = 1;
            using (var package = new ExcelPackage(redhatFile))
            {
                ExcelWorksheet rawData = package.Workbook.Worksheets[1];
                ExcelWorksheet dataSheet = package.Workbook.Worksheets.Add("RedHatData");

                int rows = rawData.Dimension.Rows;


                for (int i = 2; i <= rows; i++)
                {
                    string bulletinId = Convert.ToString(rawData.Cells[i, 1].Value);
                    string affectedPackages = dictRedHat[bulletinId];

                    // Remove tab spaces
                    affectedPackages = affectedPackages.Replace("\t", " ");


                    // Remove multiple white spaces from HTML
                    //affectedPackages = Regex.Replace(affectedPackages, "\\s+", " ");


                    affectedPackages = Regex.Replace(affectedPackages, "<[^>]*>", "");

                    List<string> parsedValues = affectedPackages.Split(new string[] { "          ", "            ", "  ", "\t\t\t", "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    parsedValues.RemoveAll(s => s.StartsWith("SHA-256: "));
                    string outputValues = ComposeRedHatOutputForExcel(parsedValues);

                    string lastAddedKey = "";
                    string lastAddedVersion = "";
                   
                    Console.WriteLine($"Working on {bulletinId}");

                    foreach (var item in parsedValues.Skip(1))
                    {


                        if (item.StartsWith("Red Hat") || item.StartsWith("JBoss Enterprise"))
                        {
                            dict.Add(item, new Dictionary<string, List<string>>());
                            lastAddedKey = item;
                        }

                        else
                        {
                            string versionName;
                            string componentName;

                            if (item.Length<10)
                            {
                                versionName = item;
                                dict[lastAddedKey].Add(item, new List<string>());
                                lastAddedVersion = item;
                            }
                            else if (item.Length>10 && !item.StartsWith("Red Hat"))
                            {
                                componentName = item;
                                dict[lastAddedKey][lastAddedVersion].Add(componentName);
                            }

                        }
                    }

                    
                    
                    foreach (var item in dict)
                    {
                        string affectedProduct = item.Key;

                        foreach (var versionComponent in item.Value)
                        {
                            int affectedComponentNumber = 1;
                            string affectedVersion = versionComponent.Key;
                            
                            foreach (var componentInfo in versionComponent.Value)
                            {
                                string affectedComponent = componentInfo;
                                string printObject = $"{bulletinId}|{affectedProduct}|{affectedVersion}|{affectedComponent}";
                                dataSheet.Cells[rowIndex++, 1].Value = printObject;
                                Console.WriteLine($"{affectedComponentNumber++}. { printObject}");
                            }
                        }
                    }
                    



                    dict.Clear();

                }

                package.Save();
                package.Dispose();
            }
        }

        private static string ComposeRedHatOutputForExcel(List<string> parsedValues)
        {
            return string.Join("|", parsedValues.Skip(1));
        }

        private static Dictionary<string, string> CreateAndFillTheValuesForSuse(FileInfo suseFile)
        {
            var dict = new Dictionary<string, string>();

            using (var package = new ExcelPackage(suseFile))
            {
                ExcelWorksheet rawData = package.Workbook.Worksheets[1];

                int rows = rawData.Dimension.Rows;

                string html = string.Empty;

                for (int i = 2; i <= rows; i++)
                {
                    string bulletinId = package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString();
                    string url = package.Workbook.Worksheets[1].Cells[i, 1].Hyperlink.ToString();


                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream))
                    {


                        html = reader.ReadToEnd();

                        HtmlDocument doc = new HtmlDocument();
                        doc.LoadHtml(html);
                        var textString = doc.DocumentNode.InnerHtml;

                        int startIndex = textString.IndexOf("Package List:");
                        int endIndex = textString.LastIndexOf("References:");
                        string coreString = textString.Substring(startIndex, endIndex - startIndex);


                        dict.Add(bulletinId, coreString);

                    }


                }
                package.Save();
                package.Dispose();

                return dict;
            }
        }

        private static Dictionary<string, string> CreateAndFillTHeValuesForRedHat(FileInfo redhatFile)
        {
            var dict = new Dictionary<string, string>();

            using (var package = new ExcelPackage(redhatFile))
            {
                ExcelWorksheet rawData = package.Workbook.Worksheets[1];

                int rows = rawData.Dimension.Rows;

                string html = string.Empty;

                for (int i = 2; i <= rows; i++)
                {
                    string bulletinId = package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString();
                    string url = package.Workbook.Worksheets[1].Cells[i, 1].Hyperlink.ToString();


                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream))
                    {


                        html = reader.ReadToEnd();

                        HtmlDocument doc = new HtmlDocument();
                        doc.LoadHtml(html);
                        var textString = doc.DocumentNode.InnerHtml;

                        int startIndex = textString.IndexOf(@"Click a package name for more details");
                        int endIndex = textString.IndexOf(@"The Red Hat security contact is ");
                        string coreString = textString.Substring(startIndex, endIndex - startIndex);


                        dict.Add(bulletinId, coreString);

                    }


                }
                package.Save();
                package.Dispose();
            }
            return dict;
        }

        private static string HtmlToPlainText(string html)
        {
            const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";//matches one or more (white space or line breaks) between '>' and '<'
            const string stripFormatting = @"<[^>]*(>|$)";//match any character between '<' and '>', even when end tag is missing
            const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";//matches: <br>,<br/>,<br />,<BR>,<BR/>,<BR />
            var lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
            var stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
            var tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);

            var text = html;
            //Decode html specific characters
            text = System.Net.WebUtility.HtmlDecode(text);
            //Remove tag whitespace/line breaks
            text = tagWhiteSpaceRegex.Replace(text, "><");
            //Replace <br /> with line breaks
            text = lineBreakRegex.Replace(text, Environment.NewLine);
            //Strip formatting
            text = stripFormattingRegex.Replace(text, string.Empty);

            return text;
        }
    }
}
