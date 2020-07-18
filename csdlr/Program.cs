using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace csdlr
{
    public class Employee
    {
        public string FirstName { get; set; }

        public void Speak()
        {
            Console.WriteLine("Hi, my name is {0}", FirstName);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            _ = args;


            //TestReflection();

            //WriteSeparator();

            //TestDynamic();

            //WriteSeparator();

            TestExcel();


            Console.ReadLine();
        }

        private static object GetASpeaker()
        {
            return new Employee { FirstName = "Scott" };
        }

        private static void TestReflection()
        {
            object o = GetASpeaker();

            o.GetType().GetMethod("Speak").Invoke(o, null);
        }

        private static void WriteSeparator()
        {
            Console.WriteLine("---");
        }

        private static void TestDynamic()
        {
            dynamic o = GetASpeaker();

            o.Speak();


            Type dogType = Assembly.Load("Dogs").GetType("Dog");

            dynamic dog = Activator.CreateInstance(dogType);

            dog.Speak();
        }

        private static void TestExcel()
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");

            dynamic excel = Activator.CreateInstance(excelType);

            excel.Visible = true;

            excel.Workbooks.Add();


            dynamic sheet = excel.ActiveSheet;

            Process[] processes = Process.GetProcesses();


            for (int i = 0; i < processes.Length; i++)
            {
                sheet.Cells[i + 1, "A"] = processes[i].ProcessName;

                sheet.Cells[i + 1, "B"] = processes[i].Threads.Count;
            }
        }
    }
}
