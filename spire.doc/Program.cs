using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spire.doc
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch timerTotal = new Stopwatch();
            timerTotal.Start();
            try
            {
                using (Document document = new Document())
                {
                    Stopwatch timer = new Stopwatch();
                    
                    
                    timer.Start();
                    document.LoadFromFile(@"C:\Users\hector\Documents\visual studio 2013\Projects\spire.doc\spire.doc\bin\Debug\doc.docx");
                    Console.WriteLine("open file: " + timer.ElapsedMilliseconds);
                    
                    timer.Restart();
                    document.Replace("##lala##", "garbage..!!!#$%&()", true, true);
                    Console.WriteLine("replace: " + timer.ElapsedMilliseconds);


                    timer.Restart();
                    document.SaveToFile("Sample.pdf", FileFormat.PDF);
                    Console.WriteLine("save PDF file: " + timer.ElapsedMilliseconds);

                }                

            }
            catch (Exception x)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine(x.Message);
            }
            Console.WriteLine("");
            Console.WriteLine("total time: " + timerTotal.ElapsedMilliseconds);
            Console.WriteLine("done...");
            Console.ReadKey();
        }
    }
}
