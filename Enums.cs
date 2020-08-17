using System.Collections.Generic;

namespace DraftPick
{
    enum Position
    {
        FLEX,
        RB,
        WR,
        TE,
        QB,
        K,
        DEF,
        
    }
    static class Ext
    {
        static internal Dictionary<Position, int> counts = new Dictionary<Position, int>()
        {
            {Position.QB, 1},
            {Position.RB, 2},
            {Position.WR, 2},
            {Position.FLEX, 6},
            {Position.K, 1},
            {Position.DEF, 1},
            {Position.TE, 1}
        };
    }
}
