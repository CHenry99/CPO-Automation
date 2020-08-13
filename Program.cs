using System;
using System.ComponentModel;
using System.IO;
using System.IO.Enumeration;
using System.Reflection;
using System.Transactions;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace Basic_Practice
{
    class Program
    {
        static void Main(string[] args)
        {
            // User Interface variables
            string CPOFile = "";
            string selection = "";
            
            // Strings to Hold Search Values
            string CPOBudget = "";
            string PEO = "";
            string date = "";
            string[] SOCOMPM = new string[3] {"","",""};
            string[] SDWXPM = new string[3] {"","",""};
            string[] SOCOMTPOC = new string[3] {"","",""};

            // Booleans to Hold "if found"
            bool foundBudget = false;
            bool foundPEO = false;
            bool foundPeriodPerformance = false;
            bool foundSOCOMPM = false;
            bool foundSDWXPM = false;
            bool foundSOCOMTPOC = false;

            // Practice Getting Full Path

            do
            {
                Console.WriteLine("1 -- Enter File Name (Must place file in the netcore3.1 Folder of the Application)");
                Console.WriteLine("2 -- Enter Full Path");
                Console.WriteLine("Select an Option:");
                selection = Console.ReadLine();

                if (selection == "1")
                {
                    Console.WriteLine("Enter File Name: ");
                    CPOFile = Console.ReadLine();
                    CPOFile = Path.GetFullPath(CPOFile, Environment.CurrentDirectory);
                }
                else if (selection == "2")
                {
                    Console.WriteLine("Enter Full Path: ");
                    CPOFile = Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("Please enter an applicable option.");
                }

            } while (selection != "1" && selection != "2");
            

           
            // Declare Application and Document to access Word File
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();

            // Declare Application and Document to access Excel File\
            Excel.Application ExApp = new Excel.Application();
            Excel.Workbook wb = ExApp.Workbooks.Add();
            Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets.Add();



            // CPOFile = @"C:\Users\Carolyn Henry\Documents\SOFWERX\CPO Baseline Practice\Basic Practice\Resources\CPO 40_H9240520F0025_PEO-SR RIPS - FE.docx";
            // Console.WriteLine(CPOFile);

            
            // If the file exists, the program can FIND the file
            if (System.IO.File.Exists(CPOFile))
            {
                Console.WriteLine("\nAccessed File");
                Console.WriteLine("Information Found: ");

                // Open the CPO File Internally
                doc = app.Documents.Open(CPOFile, ReadOnly: true);

                // Add for testing purposes in case of errors - writes document to text file to evaluate paragraph config
                // System.IO.StreamWriter sw = System.IO.File.CreateText(@"C:\Users\Carolyn Henry\Documents\SOFWERX\CPO Baseline Practice\Basic Practice\Resources\Testing Output.txt");
 
                // Iterate through all the paragraphs/lines - doc.Paragraphs.Count is the Total # of Paragraphs
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                    {

                    // sw.WriteLine("{0}", doc.Paragraphs[i + 1].Range.Text);

                        if(!foundBudget)
                        {
                             foundBudget = GetCPOValue(i + 1, ref doc, ref CPOBudget);
                        }
                    
                        if(!foundPEO)
                        {
                            foundPEO = GetPEO(i + 1, ref doc, ref PEO);
                        }

                        if(!foundPeriodPerformance) 
                        {
                            foundPeriodPerformance = GetCPODates(i + 1, ref doc, ref date);
                        }
                        
                        if (!foundSOCOMPM)
                        {
                            foundSOCOMPM = GetPOC(i + 1, ref doc, "USSOCOM AA PM", ref SOCOMPM);
                        }

                        if(!foundSDWXPM)
                        {
                            foundSDWXPM = GetSWXPM(i + 1, ref doc, ref SDWXPM);
                        }

                        if(!foundSOCOMTPOC)
                        {
                            foundSOCOMTPOC = GetPOC(i + 1, ref doc, "USSOCOM TPOC", ref SOCOMTPOC);
                        }
                        
                }

            }
            else
            {
                Console.WriteLine("Cannot Access\n");
            }

            // Output Values
            if(foundBudget)
            {
                Console.WriteLine("Budget: {0}", CPOBudget);
            }
            if(foundPEO)
            {
                Console.WriteLine("PEO: {0}", PEO);
            }
            if(foundPeriodPerformance)
            {
                Console.WriteLine("Period of Performance: {0}", date);
            }
            if(foundSDWXPM)
            {
                Console.WriteLine("DWX/SWX PM: {0}", SDWXPM[0]);
            }
            if(foundSOCOMPM)
            {
                Console.WriteLine("USSOCOM AA PM: {0}", SOCOMPM[0]);
            }
            if(foundSOCOMTPOC)
            {
                Console.WriteLine("USSOCOM TPOC PM: {0}", SOCOMTPOC[0]);
            }

            // Close the word File
            // Close the word Application
            doc.Close();
            app.Quit(false);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            

           

            // Add header values to cells

            ws.Cells[1, 1] = "Budget";
            ws.Cells[2, 1] = "Period of Performance";
            ws.Cells[3, 1] = "PEO";
            ws.Cells[4, 1] = "USSOCOM AA PM";
            ws.Cells[5, 1] = "SWX/DWX PM";
            ws.Cells[6, 1] = "USSOCOM TPOC";

            // Add values to cells
            ws.Cells[1, 2] = CPOBudget;
            ws.Cells[2, 2] = date;
            ws.Cells[3, 2] = PEO;
            ws.Cells[4, 2] = SOCOMPM[0]; ws.Cells[4, 3] = SOCOMPM[1]; ws.Cells[4, 4] = SOCOMPM[2];

            ws.Cells[5, 2] = SDWXPM[0]; ws.Cells[5, 3] = SDWXPM[1]; ws.Cells[5, 4] = SDWXPM[2];

            ws.Cells[6, 2] = SOCOMTPOC[0]; ws.Cells[6, 3] = SOCOMTPOC[1]; ws.Cells[6, 4] = SOCOMTPOC[2];

            // Save the excel workbook
            // Enter Name 

            Console.WriteLine("Enter Desired Output File Name (w/o extension): ");
            CPOFile = Console.ReadLine();

            wb.SaveAs(CPOFile, Excel.XlFileFormat.xlWorkbookNormal);

            Console.WriteLine("File Saved as: ");
            Console.WriteLine("{0}", Path.GetFullPath(CPOFile));

            //Close the Excel File
            wb.Close();
            ExApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            

            // End main
            
        }


        // Finds the Contact Information for the Specified position
        /* Line before title = PM Name
         * Line after title = PM email 
         * 2 Lines after titel = PM phone
         */
        static bool GetPOC(int currPar, ref Word.Document doc, string position, ref string[] PMDetails)
        {
            string temp = doc.Paragraphs[currPar].Range.Text;

            int distanceToEnd = doc.Paragraphs.Count - currPar;

            // Previous Solution
              
            if(temp.Contains("POINTS OF CONTACT"))
            {
                for(int i = 0; i < distanceToEnd; i++)
                {

                    if(doc.Paragraphs[currPar + i].Range.Text.Contains(position))
                    {
                        PMDetails[0] = doc.Paragraphs[currPar + i - 1].Range.Text;
                        PMDetails[1] = doc.Paragraphs[currPar + i + 1].Range.Text;
                        PMDetails[2] = doc.Paragraphs[currPar + i + 2].Range.Text;

                        return true;
                    }
                }
               
                
            } 

            return false;
        }

        static bool GetSWXPM(int currPar, ref Word.Document doc, ref string[] PMDetails)
        {
            string temp = doc.Paragraphs[currPar].Range.Text;

            int distanceToEnd = doc.Paragraphs.Count - currPar;

            // Previous Solution

            if (temp.Contains("POINTS OF CONTACT"))
            {
                for (int i = 0; i < distanceToEnd; i++)
                {

                    if (doc.Paragraphs[currPar + i].Range.Text.Contains("SWX/DWX PM") || doc.Paragraphs[currPar + i].Range.Text.Contains("SOFWERX PM") || doc.Paragraphs[currPar + i].Range.Text.Contains("DEFENSEWERX PM"))
                    {
                        PMDetails[0] = doc.Paragraphs[currPar + i - 1].Range.Text;
                        PMDetails[1] = doc.Paragraphs[currPar + i + 1].Range.Text;
                        PMDetails[2] = doc.Paragraphs[currPar + i + 2].Range.Text;

                        return true;
                    }
                }
            }

            return false;
        }

            // Find CPO Cost
            /* The parameters are the current paragraph number, a reference/pointer to the string holding the
             * the total value of the CPO, and the open word document.
             * The function will return a bool indicating if the key word is found.
             * If the keyword is found, the following paragraph (which holds the CPO value) will be placed into "value".
             */
            static bool GetCPOValue(int currPar, ref Word.Document doc, ref string value)
        {
            string temp = doc.Paragraphs[currPar].Range.Text.Trim();
            if (temp.Contains("Total Value of this Action"))
            {
                value = doc.Paragraphs[currPar + 1].Range.Text;
                return true;
            }

            return false;
        }

        // Find Executive Office
        /* Find the first instance of the phrase "(PEO)" which indicates the first time the 
         * word is used. Then eliminate the unnecessary text and retain the acronym of the
         * the executive office.  
         */
        static bool GetPEO(int currPar, ref Word.Document doc, ref string PEO)
        {            

            PEO = doc.Paragraphs[currPar].Range.Text;
            if(PEO.Contains("Program Executive Office"))
            {
                int startIndex;
                int endIndex;

                // Find where the PEO is first listed
                startIndex = PEO.IndexOf("(PEO)");

                // Eliminate text before the Into of the PEO
                PEO = PEO.Substring(startIndex, PEO.Length - startIndex);

                // Find the Acronym for PEO by seeking the next set of Brackets ( )
                startIndex = PEO.IndexOf('(',PEO.IndexOf('(') + 1 );
                endIndex = PEO.IndexOf(')');
                
                // Reduce the PEO string to only the text inside the brackets
                PEO = PEO.Substring(startIndex, endIndex);

                PEO = PEO.Remove(0, 1);
                PEO = PEO.Remove(PEO.Length - 1, 1);

                return true;
            }

            return false;
        }

        // Find CPO Dates - Start & End
        /* 
         *  NOTE - THIS IS HARD CODED BY THE TABLE, WILL NEED TO BE CHANGED.
         */
        static bool GetCPODates(int currPar, ref Word.Document doc, ref string date)
        {
            date = doc.Paragraphs[currPar].Range.Text;
            
            if(date.Contains("Performance Period") || date.Contains("PERFORMANCE PERIOD") )
            {
                date = doc.Paragraphs[currPar + 6].Range.Text;            

                return true;
            }

            return false;
        }
    }

}
