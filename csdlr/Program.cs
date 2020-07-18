using System;
using System.Collections.Generic;
using System.Linq;
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

            TestDynamic();


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
        }
    }
}
