using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace DraftPick
{
    internal class Records
    {
        static internal List<Records> Lines = new List<Records>();
        public string PlayerName;
        public Position Position;
        public string Team;
        public bool IsDrafted;
        public int Adp;
        public int Tier;
        public double[] Weeks = new double[16];

        public Records(
            string playerName,
            Position position,
            string team,
            bool isDrafted,
            int adp,
            int tier,
            double week1,
            double week2,
            double week3,
            double week4,
            double week5,
            double week6,
            double week7,
            double week8,
            double week9,
            double week10,
            double week11,
            double week12,
            double week13,
            double week14,
            double week15,
            double week16)
        {
            PlayerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            Position = position;
            Team = team;
            IsDrafted = isDrafted;
            Adp = adp;
            Tier = tier;
            int i = 0;
            Weeks[i++] = week1;
            Weeks[i++] = week2;
            Weeks[i++] = week3;
            Weeks[i++] = week4;
            Weeks[i++] = week5;
            Weeks[i++] = week6;
            Weeks[i++] = week7;
            Weeks[i++] = week8;
            Weeks[i++] = week9;
            Weeks[i++] = week10;
            Weeks[i++] = week11;
            Weeks[i++] = week12;
            Weeks[i++] = week13;
            Weeks[i++] = week14;
            Weeks[i++] = week15;
            Weeks[i++] = week16;
        }

        internal static int ReadExcel(Workbook xlWorkBook)
        {
            int pick = 0;
            _Worksheet sheet = xlWorkBook.Sheets["Weeks"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                if ((string)range[i, 4] == "x")
                {
                    pick++;
                    continue;
                }
                int j = 1;
                Lines.Add(new Records(
                    (string)range[i, j++],
                    (Position)Enum.Parse(typeof(Position), (string)range[i, j++]),
                    (string)range[i, j++],
                    ((string)range[i, j++]) == "m",
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++]),
                    Convert.ToDouble(range[i, j++])
                ));
            }
            return pick;
        }
    }
}