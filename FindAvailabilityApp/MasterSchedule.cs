using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindAvailabilityApp
{
    class MasterSchedule
    {
        //Each dictionary has a key (hour of day) and value (level of availability)
        public string Name { get; set; }
        public string Role { get; set; }

        public int[] Tuesday { get; set; }
        public int[] Monday { get; set; }
        public int[] Wednesday { get; set; }
        public int[] Thursday { get; set; }
        public int[] Friday { get; set; }
        public int[] Saturday { get; set; }
        public int[] Sunday { get; set; }

        public string[] TuesdayMembers { get; set; }
        public string[] MondayMembers { get; set; }
        public string[] WednesdayMembers { get; set; }
        public string[] ThursdayMembers { get; set; }
        public string[] FridayMembers { get; set; }
        public string[] SaturdayMembers { get; set; }
        public string[] SundayMembers { get; set; }

        public MasterSchedule()
        {
            // 11 because from 7 to 19 there are 12 hours
            this.Monday = new int[12];
            this.Tuesday = new int[12];
            this.Wednesday = new int[12];
            this.Thursday = new int[12];
            this.Friday = new int[12];
            this.Saturday = new int[12];
            this.Sunday = new int[12];
            this.MondayMembers = new string[12];
            this.TuesdayMembers = new string[12];
            this.WednesdayMembers = new string[12];
            this.ThursdayMembers = new string[12];
            this.FridayMembers = new string[12];
            this.SaturdayMembers = new string[12];
            this.SundayMembers = new string[12];
            for (int i= 0; i<=11; i++)
            {
                this.MondayMembers[i] = "0:";
                this.TuesdayMembers[i] = "0:";
                this.WednesdayMembers[i] = "0:";
                this.ThursdayMembers[i] = "0:";
                this.FridayMembers[i] = "0:";
                this.SaturdayMembers[i] = "0:";
                this.SundayMembers[i] = "0:";
            }

        }
    }
}
