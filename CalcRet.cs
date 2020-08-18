using Microsoft.Office.Interop.Excel;
using System;

namespace DraftPick
{
    internal class CalcRet : IComparable<CalcRet>
    {
        internal string playerName;
        internal int tier;
        internal int weeks;
        internal double points;
        internal Position position;
        internal int adp;

        public CalcRet(string playerName, int tier, int adp, int weeks, double points)
        {
            this.playerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            this.tier = tier;
            this.weeks = weeks;
            this.points = points;
            this.position = Position.None;
            this.adp = adp;
        }

        public int CompareTo(CalcRet other)
        {
            return this.playerName.CompareTo(other.playerName);
        }

        public CalcRet Add(CalcRet other)
        {
            return new CalcRet(this.playerName, this.tier, this.adp, this.weeks + other.weeks, this.points + other.points);
        }
    }
}