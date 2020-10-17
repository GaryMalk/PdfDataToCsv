using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PdfDataToCsv
{
    class Program
    {
        private static Dictionary<string, string> states   = new Dictionary<string, string>();
        private static DirectoryInfo outputDir = new DirectoryInfo(ConfigurationManager.AppSettings["output"]);
        private static DirectoryInfo intputDir = new DirectoryInfo(ConfigurationManager.AppSettings["input"]);
        private static Regex combinedNameRegex = new Regex(@"(?<output>.{1,})(?<start>\d{4})_(?<end>\d{4})\.pdf", RegexOptions.Singleline);
        private static Regex yearlyRegex = new Regex(@"(?<output>.{1,})(?<year>\d{4})\.pdf", RegexOptions.Singleline);

        static void Main(string[] args)
        {
            if (!outputDir.Exists)
            {
                outputDir.Create();
            }

            // copy templates prior to appending them with the yearly data
            DirectoryInfo templates = new DirectoryInfo("templates");
            foreach (FileInfo fileInfo in templates.EnumerateFiles())
            {
                fileInfo.CopyTo(Path.Combine(outputDir.FullName, fileInfo.Name), true);
            }

            // copy the State.csv file to the output directory
            FileInfo stateCsv = new FileInfo("State.csv");
            stateCsv.CopyTo(Path.Combine(outputDir.FullName, stateCsv.Name), true);

            // build map of State Names to StateIds
            StreamReader streamReader = new StreamReader(stateCsv.Name);
            _ = streamReader.ReadLine(); // throw away header
            string text = streamReader.ReadLine();
            while (text !=null)
            {
                var entry = text.Split(',');
                states.Add(entry[1], entry[0]); // import backwards; we will use the State Name as the key
                text = streamReader.ReadLine();
            }

            // get all PDF files
            DirectoryInfo pdfDir = new DirectoryInfo(@"pdf");
            foreach (var file in pdfDir.EnumerateFiles($"*.pdf", SearchOption.AllDirectories))
            {
                ProcessReports(file);
            }

            // additonal reports can be processed here, if desired
            if (intputDir.Exists)
            {
                foreach (var file in intputDir.EnumerateFiles($"*.pdf", SearchOption.AllDirectories))
                {
                    ProcessReports(file);
                }
            }

        }

        private static void ProcessReports(FileInfo file)
        {
            string data = GetPdfText(file.FullName);
            if (combinedNameRegex.IsMatch(file.Name))
            {
                // process combined reports
                GenerateCsvFromCombined(data, file.Name);
            }
            else if (yearlyRegex.IsMatch(file.Name))
            {
                // process yearly reports
                GenerateCsv(data, file.Name);
            }
            else
            {
                // unknown filename format
                throw new Exception("Filename did not fit expected format: " + file.Name);
            }
        }

        private static string GetPdfText(string filename)
        {
            PDDocument doc = PDDocument.load(filename);
            PDFTextStripper stripper = new PDFTextStripper();
            string text = stripper.getText(doc);
            doc.close();
            return text;
        }

        private static void GenerateCsvFromCombined(string data, string inputFile)
        {
            Match match = combinedNameRegex.Match(inputFile);
            string outputFile = Path.Combine(outputDir.FullName, match.Groups["output"].Value + ".csv");
            int start = int.Parse(match.Groups["start"].Value);
            int end = int.Parse(match.Groups["end"].Value);

            StreamWriter writer = new StreamWriter(outputFile, true);
            writer.AutoFlush = true;
            
            var reader = new StringReader(data);
            string text = reader.ReadLine();
            while (text != null)
            {
                // remove commas used as thousands separator which will throw off csv
                text = text.Replace(",", "");

                // only data rows: ignore all lines which do not begin with the name of a state
                var query = from kvp in states
                            where text.StartsWith(kvp.Key)
                            select kvp;

                // only data rows: ignore all lines which do not begin with the name of a state
                if (query.Count() > 0)
                {
                    var kvp = query.First();
                    // remove the state name and then split the rest of the row
                    string[] entry = text.Replace($"{kvp.Key} ", "").Split(' ');

                    int entryIndex = 0;
                    // these entries need to be flattened. 2012 through 2016 are on the same row for each state
                    // we want a separate line for each state/year so we can normalize the data in the db
                    for (int year = start; year <= end; year++)
                    {
                        // the order of the columns is Yes, No, Total, Missing
                        // if the Total (third column) == 0, this means there is no data
                        // for that state for this year, skip it so it doesn't block data when we have it
                        if (entry[entryIndex + 2] != "0")
                        {
                            writer.Write(kvp.Value); // stateId
                            writer.Write(",");
                            writer.Write(year);
                            writer.Write(",");

                            // we know that these three reports are all binary yes/no or male/female, plus total and missing
                            // so four columns total
                            for (int col = 0; col < 4; col++)
                            {
                                writer.Write(entry[entryIndex]);
                                if (col == 3)
                                {
                                    writer.WriteLine();
                                }
                                else
                                {
                                    writer.Write(",");
                                }

                                entryIndex++;
                            }
                        }
                        else
                        {
                            // increase by 4 because we skip the entire row
                            entryIndex += 4;
                        }
                    }
                }

                text = reader.ReadLine();
            }

            reader.Close();
            writer.Flush();
            writer.Close();
        }

        private static void GenerateCsv(string data, string inputFile)
        {
            Match match = yearlyRegex.Match(inputFile);
            string outputFile = Path.Combine(outputDir.FullName, match.Groups["output"].Value + ".csv");
            int year = int.Parse(match.Groups["year"].Value);

            StreamWriter writer = new StreamWriter(outputFile, true);
            writer.AutoFlush = true;
            var reader = new StringReader(data);
            string text = reader.ReadLine();
            while (text !=null)
            {
                // remove commas used as thousands separator which will throw off csv
                text = text.Replace(",", "");

                // only data rows: ignore all lines which do not begin with the name of a state
                var query = from kvp in states
                            where text.StartsWith(kvp.Key)
                            select kvp;
                
                if (query.Count() > 0)
                {
                    var kvp = query.First();
                    // replace state name with state id for normalized data
                    text = text.Replace(kvp.Key, $"{kvp.Value},{year}");

                    // text is stripped with spaces between entries, replace with commas for csv
                    writer.WriteLine(text.Replace(" ", ","));
                }

                text = reader.ReadLine();
            }

            reader.Close();
            writer.Flush();
            writer.Close();
        }
    }
}
