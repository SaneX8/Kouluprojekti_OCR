using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCRTesting
{
    internal class Row
    {
        public int left { get; set; } //tekstin etäisyys vasemmasta reunasta
        public int top { get; set; } //tekstin etäisyys ylhäältä
        public string content { get; set; } //tekstin sisältö
        public int x { get; set; } //sama käytännössä kuin left (aika turha)
        public int y { get; set; } //sama käytännössä kuin top (aika turha)
        public int page { get; set; } //millä sivulla teksti on
        public int row { get; set; } //millä rivillä teksti on

        public Row(int left, int top, string content, int x, int y, int page, int row)
        {
            this.left = left;
            this.top = top;
            this.content = content;
            this.y = y;
            this.x = x;
            this.page = page;
            this.row = row;
        }
    }
}
