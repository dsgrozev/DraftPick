using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DraftPick
{
    class Program
    {
        static int WEEK;
        static void Main(string[] args)
        {
            // Read excel
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks
                .Open(@"C:\FF\testNew.xlsx");
            // Read Defensive Data
            Records.ReadExcel(xlWorkBook);
            xlWorkBook.Close();
            xlApp.Quit();
            // For each position, produce top candidate
            foreach (Position p in Enum.GetValues(typeof(Position)))
            {
                List<CalcRet> players = CalculateBestPlayer(p).Values.ToList<CalcRet>();
                players.Sort(Compare);
                foreach (var player in players)
                {
                    Console.WriteLine(p.ToString() + ": " + player.playerName + ", Tier: " + player.tier + "; Weeks: " + player.weeks + ": " + player.points);
                }
                Console.WriteLine("-------------------------------------");
            }
            Console.ReadKey();
        }

        private static Dictionary<string, CalcRet> CalculateBestPlayer(Position p)
        {
            List<Records> records;
            Dictionary<string, CalcRet> ret = new Dictionary<string, CalcRet>();
            int minTier = 0;
            do
            {
                minTier++;
                if (p != Position.FLEX)
                {
                    records = Records.Lines.Where(x => x.Position == p).ToList();
                }
                else
                {
                    records = Records.Lines.Where(x =>
                        x.Position == Position.RB || x.Position == Position.TE || x.Position == Position.WR).ToList();
                }
                records = records.Where(x => x.Tier <= minTier).ToList();
            } while (records.Count <= Ext.counts[p]);

            int baseRecord = Ext.counts[p];

            for (int i = 0; i < 16; i++)
            {
                WEEK = i;
                records.Sort(CompareByWeek);
                for (int j = 0; j < baseRecord; j++)
                {
                    if (!records[j].IsDrafted)
                    {
                        CalcRet cr =
                            new CalcRet(records[j].PlayerName, records[j].Tier, 1, records[j].Weeks[i] - records[baseRecord].Weeks[i]);

                        if (WEEK == 0 || WEEK == 1)
                        {
                            cr.weeks = 4;
                            cr.points *= 4;
                        }
                        if (WEEK == 2 || WEEK == 3)
                        {
                            cr.weeks = 2;
                            cr.points *= 2;
                        }

                        if (!ret.ContainsKey(cr.playerName))
                        {
                            ret.Add(cr.playerName, cr);
                        }
                        else
                        {
                            ret[cr.playerName] = ret[cr.playerName].Add(cr);
                        }
                    }
                }
            }
            return ret;
        }
        private static int Compare(CalcRet x, CalcRet y)
        {
            if (x.tier == y.tier)
            {
                if (x.weeks == y.weeks)
                {
                    return -1 * x.points.CompareTo(y.points);
                }
                return -1 * x.weeks.CompareTo(y.weeks);
            }
            return x.tier.CompareTo(y.tier);
        }
        private static int CompareByWeek(Records x, Records y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            else
            {
                if (y == null)
                {
                    return -1;
                }
                else
                {
                    return -1 * x.Weeks[WEEK].CompareTo(y.Weeks[WEEK]);
                }
            }
        }
    }
}
