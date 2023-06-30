using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class DrawHelper
    {
        private Int32 currentDay;
        private Int32 currentMonth;
        private Int32 currentYear;
        private Int32 currentSpan;
        public DrawHelper(Int32 startYear, Int32 startMonth, Int32 startDay, Int32 currentSpan)
        {
            this.currentDay = startDay;
            this.currentMonth = startMonth;
            this.currentYear = startYear;
            this.currentSpan = currentSpan;

        }

        public void drawPicture(Graphics g, MyFunPictureBox picture)
        {

        }
    }


}
