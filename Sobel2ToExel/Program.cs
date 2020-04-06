using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Sobel2ToExel
{
    class Program
    {
        static int startG,stopG,startR,stopR;

        static void Main(string[] args)
        {
            //intro words
            Console.ForegroundColor = ConsoleColor.DarkGray;

            Console.WriteLine(@" Sobel2ToExel - software for data analysis.
 Copyright (C) 2019  Georgi Danovski.

 This program is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 This program is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.If not, see<http://www.gnu.org/licenses/>.
-------------------------------------------------------------------");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(" >>> Hello!");
            Console.WriteLine(" >>> Press any key to continue...");
            Console.WriteLine("");
            Console.ReadKey();

            bool repeat = true;

            while (repeat)
            {
                Body();

                Console.ForegroundColor = ConsoleColor.White;

                Console.WriteLine(" >>> Type Y to exit the program or press any kay to continue...");

                if (Console.ReadKey().Key.ToString().ToUpper() == "y".ToUpper()) repeat = false;
            }

        }
        private static void Body()
        {
            string input;

            //input directory
            do
            {
                Console.WriteLine(" >>> Input directory:");
                Console.ForegroundColor = ConsoleColor.Green;
                input = Console.ReadLine();

                if (!Directory.Exists(input))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" >>> Incorrect directory!");
                }

                Console.ForegroundColor = ConsoleColor.White;
            }
            while (!Directory.Exists(input));

            startG = GetValue("Normalize - start G:");
            stopG = GetValue("Normalize - stop G:");
            startR = GetValue("Normalize - start R:");
            stopR = GetValue("Normalize - stop R:");
            
            List<string> files = GetFiles(input);
            foreach (string str in files)
                CalculateFile(str);
        }
        private static int GetValue(string name)
        {
            int result = 0;
            string str = "";
            do
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(" >>> "+name+":");
                Console.ForegroundColor = ConsoleColor.Green;
                str = Console.ReadLine();

                if (!int.TryParse(str, out result))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" >>> Incorrect Value!");
                }

                Console.ForegroundColor = ConsoleColor.White;
            }
            while (!int.TryParse(str,out result));

            Console.ForegroundColor = ConsoleColor.White;

            return result;
        }
        private static void CalculateFile(string input)
        {
            string output, str;
            List<List<string>> data = new List<List<string>>();
            int rowL, columnL, rowLbezT;
            //output
            output = input.Substring(0, input.LastIndexOf(".")) + "_Norm.txt";
            //read file
            using (StreamReader sr = new StreamReader(input))
            {
                sr.ReadLine();//Results data type row
                data.Add(sr.ReadLine().Split(new string[] { "\t" }, StringSplitOptions.None).ToList());//titles
                sr.ReadLine();//coments

                str = sr.ReadLine();
                while (str != null)
                {
                    data.Add(str.Split(new string[] { "\t" }, StringSplitOptions.None).ToList());//values
                    str = sr.ReadLine();
                }

                int Length = (data[0].Count() - 6) / 2 + 1;

                for (int i = 0; i < data.Count(); i++)
                {
                    data[i] = data[i].GetRange(0, Length);
                }

                data[0][0] = "T(sec.)";
            }
            //calculations
            columnL = data.Count();
            rowL = data[0].Count();
            rowLbezT = rowL - 1;

            //find max
            for (int col = 1; col < rowL; col++)
            {
                data[0].Add("Max_" + data[0][col]);
                data[1].Add("=MAX(" + ColumnLabel(col + 1+2*rowLbezT) + "2:" + ColumnLabel(col + 1 + 2 * rowLbezT) + columnL + ")");
                for (int row = 2; row < columnL; row++)
                {
                    data[row].Add("");
                }
            }
            //find To0
            for (int col = 1; col < rowL; col++)
            {
                data[0].Add("To0_" + data[0][col]);
                for (int row = 1; row < columnL; row++)
                    if (input.EndsWith("_G.txt"))
                    {
                        // data[row].Add("=" + ColumnLabel(col + 1) + (row + 1).ToString() + "-AVERAGE(" + ColumnLabel(col + 1) + "$2:" + ColumnLabel(col + 1) + "$32)");
                        data[row].Add("=" + ColumnLabel(col + 1) + (row + 1).ToString() + "-AVERAGE(" + ColumnLabel(col + 1) + "$"+startG+":" + ColumnLabel(col + 1) + "$"+stopG+")");
                    }
                    else
                    {
                        // data[row].Add("=" + ColumnLabel(col + 1) + (row + 1).ToString() + "-AVERAGE(" + ColumnLabel(col + 1) + "$62:" + ColumnLabel(col + 1) + "$142)");
                        data[row].Add("=" + ColumnLabel(col + 1) + (row + 1).ToString() + "-AVERAGE(" + ColumnLabel(col + 1) + "$"+startR+":" + ColumnLabel(col + 1) + "$"+stopR+")");
                    }
            }

            //find nTo0
            for (int col = 1; col < rowL; col++)
            {
                data[0].Add("nTo0_" + data[0][col]);

                int To0col = rowL + rowLbezT + col;//To0 cell
                string max = ColumnLabel(To0col - rowLbezT) + "$2";//max cell

                for (int row = 1; row < columnL; row++)
                {
                    data[row].Add("=" + ColumnLabel(To0col) + (row + 1).ToString() + "/" + max);
                }
            }
            //StDev and avg
            {
                int To0colStart = rowL + rowLbezT + 1;
                int To0colStop = To0colStart + rowLbezT-1;
                int nTo0colStart = To0colStop + 1;
                int nTo0colStop = nTo0colStart + rowLbezT-1;
                int avgCol = data[0].Count() + 1;

                data[0].AddRange(new string[] { "nAvgMob", "nnAvgMob", "nStDevMob", "AvgMob", "StDevMob" });

                for (int row = 1; row < columnL; row++)
                {
                    data[row].AddRange(new string[] {
                        "=AVERAGE(" + ColumnLabel(nTo0colStart) + (row+1).ToString() + ":" + ColumnLabel(nTo0colStop) + (row+1).ToString()+")",
                        "=" + ColumnLabel(avgCol) + (row+1).ToString() + "/MAX(" + ColumnLabel(avgCol) + "$2:" + ColumnLabel(avgCol) + "$" + columnL + ")",
                        "=STDEV.S(" + ColumnLabel(nTo0colStart) + (row+1).ToString() + ":" + ColumnLabel(nTo0colStop) + (row+1).ToString()+")",
                        "=AVERAGE(" + ColumnLabel(To0colStart) + (row+1).ToString() + ":" + ColumnLabel(To0colStop) + (row+1).ToString()+")",
                        "=STDEV.S(" + ColumnLabel(To0colStart) + (row+1).ToString() + ":" + ColumnLabel(To0colStop) + (row+1).ToString()+")"
                    });
                }
            }

            string[] results = new string[data.Count()];

            for (int i = 0; i < data.Count(); i++)
            {
                results[i] = string.Join("\t", data[i]);
            }
            try
            {
                File.WriteAllLines(output, results);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(" >>> Output file is opened in other program!");
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
        public static string ColumnLabel(int col)
        {
            var dividend = col;
            var columnLabel = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnLabel = Convert.ToChar(65 + modulo).ToString() + columnLabel;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnLabel;
        }
        public static int ColumnIndex(string colLabel)
        {
            // "AD" (1 * 26^1) + (4 * 26^0) ...
            var colIndex = 0;
            for (int ind = 0, pow = colLabel.Count() - 1; ind < colLabel.Count(); ++ind, --pow)
            {
                var cVal = Convert.ToInt32(colLabel[ind]) - 64; //col A is index 1
                colIndex += cVal * ((int)Math.Pow(26, pow));
            }
            return colIndex;
        }
        private static List<string> GetFiles(string dir)
        {
            List<string> result = new List<string>();

            DirectoryInfo di = new DirectoryInfo(dir);

            foreach (var fi in di.GetFiles())
                if (fi.Extension == ".txt")
                    result.Add(fi.FullName);

            di = null;

            return result;
        }
    }
}

