using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;  


namespace FindAvailabilityApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string root = Path.GetFullPath(args[0]).Replace(@"\", @"\\");

            String[] files = Directory.GetFiles(root + @"\\", "AvailabilityDoc**.xlsx");
            List<Member> members = new List<Member>();
            foreach (string file in files)
            { 
                Console.WriteLine(file);
                Member member = loadMemberSchedule(file);
                members.Add(member);
            }
            //the output object
            MasterSchedule masterSchedule = new MasterSchedule();
            masterSchedule.Name = "GroupSchedule";
            
            //doc comparison logic
            //for each scrum master
            foreach (var member in members)
            {
                if (member.Role.Contains("Scrum Master"))
                {
                    for (int hour = 0; hour <= 11; hour++)
                    {
                        //
                        if (member.Monday[hour] >0)
                        {
                            masterSchedule.Monday[hour] = member.Monday[hour];
                            masterSchedule.MondayMembers[hour] = (iterateMemberCount(masterSchedule.MondayMembers[hour])+ member.Name+" , ");
                        }
                        if (member.Tuesday[hour] > 0)
                        {
                            masterSchedule.Tuesday[hour] = member.Tuesday[hour];
                            masterSchedule.TuesdayMembers[hour] = (iterateMemberCount(masterSchedule.TuesdayMembers[hour]) + member.Name + " , ");
                        }
                        if (member.Wednesday[hour] > 0)
                        {
                            masterSchedule.Wednesday[hour] = member.Wednesday[hour];
                            masterSchedule.WednesdayMembers[hour] = (iterateMemberCount(masterSchedule.WednesdayMembers[hour]) + member.Name + " , ");
                        }
                        if (member.Thursday[hour] > 0)
                        {
                            masterSchedule.Thursday[hour] = member.Thursday[hour];
                            masterSchedule.ThursdayMembers[hour] = (iterateMemberCount(masterSchedule.ThursdayMembers[hour]) + member.Name + " , ");
                        }
                        if (member.Friday[hour] > 0)
                        {
                            masterSchedule.Friday[hour] = member.Friday[hour];
                            masterSchedule.FridayMembers[hour] = (iterateMemberCount(masterSchedule.FridayMembers[hour]) + member.Name + " , ");
                        }
                        if (member.Saturday[hour] > 0)
                        {
                            masterSchedule.Saturday[hour] = member.Saturday[hour];
                            masterSchedule.SaturdayMembers[hour] = (iterateMemberCount(masterSchedule.SaturdayMembers[hour]) + member.Name + " , ");
                        }
                        if (member.Sunday[hour] > 0)
                        {
                            masterSchedule.Sunday[hour] = member.Sunday[hour];
                            masterSchedule.SundayMembers[hour] = (iterateMemberCount(masterSchedule.SundayMembers[hour]) + member.Name + " , ");
                        }
                    }
                }
                
            }
            //for each development team member
            foreach (var member in members)
            {
                if (!member.Role.Contains("Scrum Master"))
                {

                    for(int hour = 0; hour<=11; hour++)
                    {
                        if (masterSchedule.Monday[hour] > 0 && member.Monday[hour] > 0)
                        {
                            masterSchedule.MondayMembers[hour] = (iterateMemberCount(masterSchedule.MondayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Tuesday[hour] > 0 && member.Tuesday[hour] > 0)
                        {
                            masterSchedule.TuesdayMembers[hour] = (iterateMemberCount(masterSchedule.TuesdayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Wednesday[hour] > 0 && member.Wednesday[hour] > 0)
                        {
                            masterSchedule.WednesdayMembers[hour] = (iterateMemberCount(masterSchedule.WednesdayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Thursday[hour] > 0 && member.Thursday[hour] > 0)
                        {
                            masterSchedule.ThursdayMembers[hour] = (iterateMemberCount(masterSchedule.ThursdayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Friday[hour] > 0 && member.Friday[hour] > 0)
                        {
                            masterSchedule.FridayMembers[hour] = (iterateMemberCount(masterSchedule.FridayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Saturday[hour] > 0 && member.Saturday[hour] > 0)
                        {
                            masterSchedule.SaturdayMembers[hour] = (iterateMemberCount(masterSchedule.SaturdayMembers[hour]) + member.Name + " , ");
                        }
                        if (masterSchedule.Sunday[hour] > 0 && member.Sunday[hour] > 0)
                        {
                            masterSchedule.SundayMembers[hour] = (iterateMemberCount(masterSchedule.SundayMembers[hour]) + member.Name + " , ");
                        }
                    }
                }
            }
            root= root.Replace(@"\\", @"\");
            createExcell(masterSchedule).ExportToExcel(root+@"\SummarisedSchedule.xlsx");
            
        }

        public static string iterateMemberCount(string members)
        {
            string[] getNum = members.Split(':');
            int count = Convert.ToInt32(getNum[0].Trim());
            count++;
            getNum[0] = count.ToString();
            return string.Join(":", getNum);
        }

        public static System.Data.DataTable createExcell(MasterSchedule masterSchedule)
        {
            string[] availabilitySignifier = new string[3];
            availabilitySignifier[0] = "NA";
            availabilitySignifier[1] = "FX";
            availabilitySignifier[2] = "AV";
            System.Data.DataTable table = new System.Data.DataTable("ParentTable");
                table.Columns.Add("", typeof(string));
                table.Columns.Add("Monday", typeof(string));
            table.Columns.Add("Monday Members", typeof(string));
            table.Columns.Add("Tuesday", typeof(string));
            table.Columns.Add("Tuesday Members", typeof(string));
            table.Columns.Add("Wednesday", typeof(string));
            table.Columns.Add("Wednesday Members", typeof(string));
            table.Columns.Add("Thursday", typeof(string));
            table.Columns.Add("Thursday Members", typeof(string));
            table.Columns.Add("Friday", typeof(string));
            table.Columns.Add("Friday Members", typeof(string));
            table.Columns.Add("Saturday", typeof(string));
            table.Columns.Add("Saturday Members", typeof(string));
            table.Columns.Add("Sunday", typeof(string));
            table.Columns.Add("Sunday Members", typeof(string));
            string[] hoursArray = new[]
            {
                "8:00", "9:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00"
            };
            
            for (int i = 0; i < 11; i++)
            {
                table.Rows.Add(hoursArray[i], availabilitySignifier[masterSchedule.Monday[i]],masterSchedule.MondayMembers[i], availabilitySignifier[masterSchedule.Tuesday[i]],masterSchedule.TuesdayMembers[i], availabilitySignifier[masterSchedule.Wednesday[i]],masterSchedule.WednesdayMembers[i], availabilitySignifier[masterSchedule.Thursday[i]],masterSchedule.ThursdayMembers[i], availabilitySignifier[masterSchedule.Friday[i]],masterSchedule.FridayMembers[i], availabilitySignifier[masterSchedule.Saturday[i]],masterSchedule.SaturdayMembers[i], availabilitySignifier[masterSchedule.Sunday[i]],masterSchedule.SundayMembers[i]);
            } 
            return table;
        }
        public static Member loadMemberSchedule(string filePath)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet1.UsedRange;


            // new Member
            Member member = new Member();
            if (xlRange.Cells[1, 3] != null && xlRange.Cells[1, 3].Value2 != null)
                member.Name = xlRange.Cells[1, 3].Value2.ToString();
            if (xlRange.Cells[2, 3] != null && xlRange.Cells[2, 3].Value2 != null)
                member.Role = xlRange.Cells[2, 3].Value2.ToString();
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= 8; i++)
            {
                int caseSwitch = i;
                string availability ="";
                switch (caseSwitch)
                {
                    case 2:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Monday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Monday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Monday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                    case 3:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Tuesday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Tuesday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Tuesday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                    case 4:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Wednesday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Wednesday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Wednesday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                    case 5:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Thursday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Thursday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Thursday[j - 6] = 2;
                                    break;
                            }  
                        }

                        break;
                    case 6:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Friday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Friday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Friday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                    case 7:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Saturday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Saturday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Saturday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                    case 8:
                        for (int j = 6; j <= 17; j++)
                        {
                            if (xlRange.Cells[j, i] != null && xlRange.Cells[j, i].Value2 != null)
                                availability = xlRange.Cells[j, i].Value2.ToString();
                            switch (availability)
                            {
                                case "NA":
                                    member.Sunday[j - 6] = 0;
                                    break;
                                case "FX":
                                    member.Sunday[j - 6] = 1;
                                    break;
                                case "AV":
                                    member.Sunday[j - 6] = 2;
                                    break;
                            }
                        }

                        break;
                }
            }

            return member;
        }
    }
}
