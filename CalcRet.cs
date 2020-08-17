using System;

namespace DraftPick
{
    internal struct CalcRet : IComparable<CalcRet>
    {
        internal string playerName;
        internal int tier;
        internal int weeks;
        internal double points;

        public CalcRet(string playerName, int tier, int weeks, double points)
        {
            this.playerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            this.tier = tier;
            this.weeks = weeks;
            this.points = points;
        }

        public int CompareTo(CalcRet other)
        {
            return this.playerName.CompareTo(other.playerName);
        }

        public CalcRet Add(CalcRet other)
        {
            return new CalcRet(this.playerName, this.tier, this.weeks + other.weeks, this.points + other.points);
        }
    }
}