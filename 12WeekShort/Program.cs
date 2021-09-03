using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel ;

namespace _12WeekShort
{
    class Program
    {
        public static int site = 10;
        static void Main(string[] args)
        {
          

            string CSVoutput = "Work Order,Parent Part,WIP QTY,Required Date,Promise Date,Revised Promise Date,Short Part,Short QTY,Short BOM Level,Short Part Description,Contracts\n";

            List<CSVLine> c = new List<CSVLine>();
            mtms.DataTable1DataTable WORD = new mtms.DataTable1DataTable();
            mtmsTableAdapters.DataTable1TableAdapter dtaWORD = new mtmsTableAdapters.DataTable1TableAdapter();

            double weeks = 1;
            weeks:
            Console.WriteLine("how many weeks");
            string weekshold =
            Console.ReadLine();
            if (!double.TryParse(weekshold, out weeks))
            {
                goto weeks;
            }
            Console.WriteLine();
            Console.WriteLine("Running Open WIP Report");
            
            dtaWORD.Fill(WORD,DateTime.Today.AddDays(7*weeks));
            mTMSLibrary.Screens m = new mTMSLibrary.Screens("10");
            int i = 2;
            int y = 1;
            Console.WriteLine("Checking all WIPs");

            foreach (mtms.DataTable1Row row in WORD)
            {
                Console.WriteLine("{0} out of {1}", y, WORD.Count());
                y++;
                List<string> hold;
                DateTime dthold = new DateTime(2020, 01, 01);
                mtms.MISC_DATADataTable MISC = new mtms.MISC_DATADataTable();
                mtmsTableAdapters.MISC_DATATableAdapter dtaMISC = new mtmsTableAdapters.MISC_DATATableAdapter();
               

                dtaMISC.Fill(MISC, site.ToString(), short.Parse(row.WORD1_REF.Substring(6, 2)), row.WORD1_REF.Substring(0, 6));
                if (MISC.Count > 0)
                {
                    hold = m.WPQ43(row.WORDPART, int.Parse(row.WORDQTY_REQ.ToString()), int.Parse(row.WORDPATH.ToString()), MISC[0].MISCDATE_3);
                    dthold = MISC[0].MISCDATE_3;
                }
                else
                {
                    hold = m.WPQ43(row.WORDPART, int.Parse(row.WORDQTY_REQ.ToString()), int.Parse(row.WORDPATH.ToString()), row.ORDSDATE_PROMISE);
                }
                if (hold.Count > 0)
                {   

                    foreach (string s in hold)
                    {
                        //Short Part Description,Contracts\n";

                        //Console.WriteLine(s);
                        var g = s.Split('~');
                        c.Add(new CSVLine()
                        {
                            workOrder = row.WORD1_REF,
                            parentPart = row.WORDPART,
                            wipQTY = double.Parse(row.WORDQTY_REQ.ToString()),
                            dateRequired = row.ORDSDATE_REQUEST,
                            dateProm = row.ORDSDATE_PROMISE,
                            dateRevProm = dthold,
                            shortPart = g[1],
                            shortQTY = double.Parse(g[3]),
                            shortBOMLevel = int.Parse(g[0]),
                            shortPartDesc = g[2].Replace(",", " "),
                            contracts = row.ORDS5_SAL_BUY,
                            buyer = ""

                        }) ;                     
                       
                        //CSVoutput += row.WORD1_REF + ",";
                        //CSVoutput +=  row.WORDPART + ",";
                        //CSVoutput += row.WORDQTY_REQ + ",";
                        //CSVoutput += row.ORDSDATE_REQUEST.ToShortDateString() + ",";
                        //CSVoutput += row.ORDSDATE_PROMISE.ToShortDateString() + ",";
                        //CSVoutput += dthold.ToShortDateString() + ",";
                        //CSVoutput += g[1] + ",";
                        //CSVoutput += g[3] + ",";
                        //CSVoutput += g[0] + ",";
                        //CSVoutput += g[2].Replace(","," ") + ",";
                        //CSVoutput += row.ORDS5_SAL_BUY;
                        //CSVoutput += "\n";
                        i++;
                    } 
                }            
            }
            getBuyers(c);
            string csvpath = 
            WriteToCSV(csvlinesToString(c));
            try
            {
                System.Diagnostics.Process.Start(csvpath);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message + '\n' + csvpath);
            }
            Console.WriteLine(WORD.Count);



        }
        public static string csvlinesToString(List<CSVLine> c)
        {
            string output = "Work Order,Parent Part,WIP QTY,Required Date,Promise Date,Revised Promise Date,Short Part,Short QTY,Short BOM Level,Short Part Description,Contracts,Buyer\n";
            foreach (CSVLine l in c)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}\n",
                    l.workOrder, l.parentPart, l.wipQTY, l.dateRequired.ToShortDateString(), l.dateProm.ToShortDateString(), l.dateRevProm.ToShortDateString(),
                    l.shortPart, l.shortQTY, l.shortBOMLevel, l.shortPartDesc, l.contracts, l.buyer);
                output +=sb.ToString() ;
            }
            return output;
        } 
       public static string SQLString(double weeks)
        {
            StringBuilder hold = new StringBuilder();
             hold.AppendFormat( "SELECT [MTMS.WORD_DATA].* , [MTMS.ORDS_DATA].* \n" +
                          "FROM MTMS.WORD_DATA, MTMS.ORDS_DATA \n" +
                          "WHERE MTMS.WORD_DATA.WORD1_CO_SITE = MTMS.ORDS_DATA.ORDS1_CO_SITE AND " +
                          "SUBSTR(MTMS.WORD_DATA.WORD1_REF, 1, 6) = SUBSTR(MTMS.ORDS_DATA.ORDS1_REF, 1, 6) AND " +
                          "TO_NUMBER(SUBSTR(MTMS.WORD_DATA.WORD1_REF, 7, 2)) = TO_NUMBER(MTMS.ORDS_DATA.ORDS1_LINE) AND " +
                          "(MTMS.WORD_DATA.WORD1_CO_SITE = 10) AND(MTMS.WORD_DATA.WORDSTATUS <= 12) AND " +
                          "(MTMS.WORD_DATA.WORDDATE_REQ <= {0})" , DateTime.Today.AddDays(7*weeks));
            return hold.ToString();
        }

        //public static mtms.PART_DATADataTable getPartTable()
        //{
        //    mtms.PART_DATADataTable PART = new mtms.PART_DATADataTable();
        //    mtmsTableAdapters.PART_DATATableAdapter dtaPART = new mtmsTableAdapters.PART_DATATableAdapter();
        //    dtaPART.SelectCommand.CommandText = getPARTSQL();
        //    dtaPART.Fill(PART);
        //    return PART;
        //}
        public static mtms.ALOC_DATADataTable getAlocationTable(mtms.WORD_DATADataTable WORD)
        {
            mtms.ALOC_DATADataTable ALOC = new mtms.ALOC_DATADataTable();
            mtmsTableAdapters.ALOC_DATATableAdapter dtaALOC = new mtmsTableAdapters.ALOC_DATATableAdapter();
            dtaALOC.SelectCommand.CommandText = getALOCSQL(WORD);
            dtaALOC.Fill(ALOC);
            return ALOC;
        }
        public static string getPARTSQL()
        {
            string hold =
            @"SELECT * FROM MTMS.PART_DATA
              WHERE PART1_CO_SITE = '" + site.ToString() + @"'";
            return hold;
        }
        public static string getALOCSQL(mtms.WORD_DATADataTable WORD)
        {
            string hold = 
            @"SELECT * FROM MTMS.ALOC_DATA
              WHERE ALOC1_CO_SITE = '" + site.ToString() + @"' AND  ALOC1_REF_LINE > 0 AND ALOC1_REF IN " + ListToSQLString( GetWIPList(WORD));
            return hold;
        } 
        public static mtms.WORD_DATADataTable getWorkOrderTable(DateTime FutureDate) 
        {
            mtms.WORD_DATADataTable WORD = new mtms.WORD_DATADataTable();
            mtmsTableAdapters.WORD_DATATableAdapter dtaWORD = new mtmsTableAdapters.WORD_DATATableAdapter();
            dtaWORD.Fill(WORD, FutureDate, site.ToString() );
            return WORD;
        }
        public static List<string> GetPartList(mtms.WORD_DATADataTable WORD)
        {
            List<string> output = new List<string>();
            foreach (mtms.WORD_DATARow row in WORD)
            {
                if (!output.Contains(row.WORDPART))
                {
                    output.Add(row.WORDPART);
                }
            }
            return output;
        }
        public static List<string> GetPartList(mtms.ALOC_DATADataTable ALOC)
        {
            List<string> output = new List<string>();
            foreach (mtms.ALOC_DATARow row in ALOC)
            {
                if (!output.Contains(row.ALOCPART))
                {
                    output.Add(row.ALOCPART);
                }
            }
            return output;
        }
        public static List<string> GetWIPList(mtms.WORD_DATADataTable WORD)
        {
            List<string> output = new List<string>();
            foreach (mtms.WORD_DATARow row in WORD)
            {
                if (!output.Contains(row.WORD1_REF))
                {
                    output.Add(row.WORD1_REF);
                }
            }
            return output;
        }
        public static string ListToSQLString (List<string> input)
        {
            string output = @"(";
            foreach(string s in input)
            {
                output += @"'" + s + @"',";
            }
            output = output.TrimEnd(',');
            output += ")";
            return output;
        }
        public static string WriteToCSV (string input)
        {
            string hold = 
            System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".csv");

            System.IO.File.WriteAllText(hold, input);

            return hold;

        }

        public static void getBuyers (List<CSVLine> lines)
        {
            int i = 1;
            Dictionary<string, string> partBuyer = new Dictionary<string, string>();
            foreach (CSVLine c in lines)
            {

                Console.WriteLine("gethering buyer for {0} - {1} out of {2}", c.shortPart, i, lines.Count());
                  i++;
                if (!partBuyer.Keys.Contains(c.shortPart))
                {
                    mtmsTableAdapters.PART_DATATableAdapter p = new mtmsTableAdapters.PART_DATATableAdapter();
                    foreach (mtms.PART_DATARow r in p.GetData(c.shortPart.PadRight(20)).Rows)
                    {
                        partBuyer.Add(c.shortPart, r.PARTBUYER);                        
                        break;
                    }
                    
                }
                c.buyer = partBuyer[c.shortPart];
            }

        }

        public class CSVLine
        {
            public string workOrder { get; set; }
            public string parentPart{ get; set; }
            public double wipQTY { get; set; }
            public DateTime dateRequired { get; set; }
            public DateTime dateProm { get; set; }
            public DateTime dateRevProm { get; set; }
            public string shortPart { get; set; }
            public double shortQTY { get; set; }
            public int shortBOMLevel { get; set; }
            public string shortPartDesc { get; set; }
            public string contracts { get; set; }

            public string buyer { get; set; }

           

        }
    }



    /// <Logic>
    /// Gather all Open WIPs and ones due in the time span
    /// Gather all Alocations For those WIPs
    /// Gather all Purchase orders For the parts in the WIPs and Allocations
    /// 
    /// Gather all parts effected by the allocations, WIPs amd POs.
    /// 
    /// Step through each day.
    /// Subtract allocated parts from that day. 
    /// check for >= 0 stock and flag fails 
    /// add WIPed part 
    /// add POed parts
    /// </Logic>



}
