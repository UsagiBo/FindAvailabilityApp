using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace FindAvailabilityApp
{
    class Member
    {

        private string[] _tuesday;

        //Eachday of the week has a dictionary 
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

        public Member()
        {
            // 11 because from 7 to 19 there are 12 hours
                this.Monday = new int[12];
                this.Tuesday = new int[12];
                this.Wednesday = new int[12];
                this.Thursday = new int[12];
                this.Friday = new int[12];
                this.Saturday = new int[12]; 
                this.Sunday = new int[12];


        }
    }
}
