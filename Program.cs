using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DraftPick
{
    class Program
    {
        static int WEEK;
        static int NUMBER_OF_PLAYERS = 10;
        static void Main(string[] args)
        {
        // Read excel
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\FF\testNew.xlsx");
            int pick = Records.ReadExcel(xlWorkBook);
            xlWorkBook.Close();
            xlApp.Quit();
            // For each position, produce top candidate
            List<CalcRet> playersTotal = new List<CalcRet>();
            foreach (Position p in Enum.GetValues(typeof(Position)))
            {
                if (p == Position.None)
                {
                    continue;
                }
                List<CalcRet> players = CalculateBestPlayer(p).Values.ToList();
                players.Sort(Compare);
                foreach (var player in players)
                {
                    Console.WriteLine(p.ToString() + ": " + player.playerName + ", ADP: " + player.adp + ", Tier: " + player.tier + "; Weeks: " + player.weeks + ": " + player.points);
                    CalcRet temp = player;
                    temp.position = p;
                    playersTotal.Add(temp);
                }
                Console.WriteLine("-------------------------------------");
            }
            pick += Records.Lines.FindAll(x => x.IsDrafted == true).Count;
            pick++;
            int nextPick = CalcNextPick(pick);
            playersTotal.Sort(Compare);
            Console.WriteLine("Pick: " + pick);
            Console.WriteLine("Next Pick: " + nextPick);
            CalcRet target = null;
            for (int i = 0; i < 10; i++)
            {
                if (i >= playersTotal.Count)
                {
                    break;
                }
                var player = playersTotal[i];
                if (player.adp < nextPick)
                {
                    Console.WriteLine(player.position + ": " + player.playerName + ", ADP: " + player.adp + ", Tier: " + player.tier + "; Weeks: " + player.weeks + ": " + player.points);
                }
                else
                {
                    if (i == 0)
                    {
                        target = player;
                    }
                }
            }
            if (target != null)
            {
                Console.WriteLine("TARGET ==> " + target.position + ": " + target.playerName + ", ADP: " + target.adp + ", Tier: " + target.tier + "; Weeks: " + target.weeks + ": " + target.points);
            }
            Console.ReadKey();
        }

        private static int CalcNextPick(int pick)
        {
            pick -= 1;
            int delta = pick % NUMBER_OF_PLAYERS;
            int diff = NUMBER_OF_PLAYERS - delta;
            return pick + diff * 2;
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
                            new CalcRet(records[j].PlayerName, records[j].Tier, records[j].Adp, 1, records[j].Weeks[i] - records[baseRecord].Weeks[i]);

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
            //if (x.tier == y.tier)
            //{
                if (x.weeks == y.weeks)
                {
                    return -1 * x.points.CompareTo(y.points);
                }
                return -1 * x.weeks.CompareTo(y.weeks);
            //}
            //return x.tier.CompareTo(y.tier);
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
