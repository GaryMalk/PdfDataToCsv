﻿using org.apache.pdfbox.pdmodel;
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
        private static Regex combinedNameRegex = new Regex(@"(?<output>.{1,})(?<start>\d{4})_(?<end>\d{4})\.pdf", RegexOptions.Singleline);

        static void Main(string[] args)
        {
            // get list of reports that are yearly
            string[] yearly = ConfigurationManager.AppSettings["yearly"].Split(';');
            
            // build map of State Names to StateIds
            StreamReader streamReader = new StreamReader("State.csv");
            _ = streamReader.ReadLine(); // throw away header
            string text = streamReader.ReadLine();
            while (text !=null)
            {
                var entry = text.Split(',');
                states.Add(entry[1], entry[0]); // import backwards; we will use the State Name as the key
                text = streamReader.ReadLine();
            }

            if (!outputDir.Exists)
            {
                outputDir.Create();
            }

            // copy templates prior to appending them with the yearly data
            DirectoryInfo templates = new DirectoryInfo("acf data\\templates");
            foreach (FileInfo fileInfo in templates.EnumerateFiles())
            {
                fileInfo.CopyTo(Path.Combine(outputDir.FullName, fileInfo.Name), true);
            }

            // process the yearly data
            DirectoryInfo directory = new DirectoryInfo(@"acf data");
            var subDirs = directory.EnumerateDirectories("20??");
            foreach (string pattern in yearly)
            {
                string outputFile = Path.Combine(outputDir.FullName, pattern + ".csv");
                foreach (var subDir in subDirs)
                {
                    foreach (var file in subDir.EnumerateFiles($"{pattern}20??.pdf"))
                    {
                        string data = GetPdfText(file.FullName);
                        // directory name is the year
                        GenerateCsv(data, outputFile, file.Directory.Name);
                    }
                }
            }

            DirectoryInfo combined = new DirectoryInfo("acf data\\Combined");
            foreach (var file in combined.EnumerateFiles("*.pdf"))
            {
                string data = GetPdfText(file.FullName);
                GenerateCsvFromCombined(data, file.Name);
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
            if (!match.Success)
            {
                throw new Exception("Filename did not fit expected format: " + inputFile);
            }

            string outputFile = Path.Combine(outputDir.FullName, match.Groups["output"].Value + ".csv");
            int start = int.Parse(match.Groups["start"].Value);
            int end = int.Parse(match.Groups["end"].Value);

            StreamWriter writer = new StreamWriter(outputFile);
            writer.AutoFlush = true;
            
            // we know what the headers should be so write them
            if (inputFile.Contains("gender"))
            {
                writer.WriteLine("StateId,Year,Male,Female,Total,Missing");
            }
            else
            {
                writer.WriteLine("StateId,Year,Yes,No,Total,Missing");
            }

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

        private static void GenerateCsv(string data, string outputFile, string year)
        {
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
