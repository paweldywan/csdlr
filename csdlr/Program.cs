using IronRuby;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

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


            TestReflection();

            WriteSeparator();

            TestDynamic();

            WriteSeparator();

            TestExcel();

            WriteSeparator();

            TestExpandoObject();

            WriteSeparator();

            GenerateXML();

            WriteSeparator();

            DumpEmployees();

            WriteSeparator();

            TestDynamicObject();

            WriteSeparator();

            TestIronRuby();


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

        private static void TestExpandoObject()
        {
            dynamic expando = new ExpandoObject();

            expando.Name = "Scott";

            expando.Speak = new Action(() => Console.WriteLine(expando.Name));


            expando.Speak();


            foreach (var memeber in expando)
            {
                Console.WriteLine(memeber.Key);
            }
        }

        private static void GenerateXML()
        {
            var employees = new[] { "Scott", "Poonam", "Paul", "Karthik", "Nirhir" };

            XDocument xDocument = new XDocument(
                    new XElement("Employees",
                        employees.Select(e =>
                            new XElement("Employee",
                                new XElement("FirstName", e)
                            ))
                    ));

            xDocument.Save("Employees.xml");
        }

        private static void DumpEmployees()
        {
            var doc = XDocument.Load("Employees.xml");

            foreach (var element in doc.Element("Employees")
                                       .Elements("Employee"))
            {
                Console.WriteLine(element.Element("FirstName").Value);
            }


            var doc2 = doc.AsExpando();

            foreach (var employee in doc2.Employees)
            {
                Console.WriteLine(employee.FirstName);
            }


            var doc3 = doc.ToExpando();

            foreach (var employee in doc3.Employees.Employee)
            {
                Console.WriteLine(employee.FirstName);
            }


            var doc4 = doc.Convert();

            foreach (var employee in doc4.Elements)
            {
                foreach (var firstName in employee.Elements)
                {
                    Console.WriteLine(firstName.Value);
                }
            }


            var doc5 = doc.AsExpando2();

            foreach (var employee in doc5.Employees)
            {
                foreach (var firstName in employee.Employee)
                {
                    Console.WriteLine(firstName.FirstName);
                };
            }


            var doc6 = doc.Parse();

            foreach (var employee in doc6.Employees.Employee)
            {
                Console.WriteLine(employee.FirstName);
            }
        }

        private static void TestDynamicObject()
        {
            dynamic doc = new DynamicXml("Employees.xml");

            foreach (var employee in doc.Employees)
            {
                Console.WriteLine(employee.FirstName);
            }
        }

        private static void TestIronRuby()
        {
            var engine = Ruby.CreateEngine();

            var scope = engine.CreateScope();

            scope.SetVariable("employee", new Employee { FirstName = "Scott C#" });

            engine.ExecuteFile("Program.rb", scope);


            dynamic ruby = engine.Runtime.Globals;

            dynamic person = ruby.Person.@new();

            person.firstName = "Scott Ruby";

            person.speak();
        }
    }

    internal class DynamicXml : DynamicObject, IEnumerable
    {
        private readonly dynamic _xml;

        public DynamicXml(string fileName)
        {
            _xml = XDocument.Load(fileName);
        }

        public DynamicXml(dynamic xml)
        {
            _xml = xml;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            var xml = _xml.Element(binder.Name);

            if (xml != null)
            {
                result = new DynamicXml(xml);

                return true;
            }

            result = null;

            return false;
        }

        public IEnumerator GetEnumerator()
        {
            foreach (var child in _xml.Elements())
            {
                yield return new DynamicXml(child);
            }
        }

        public static implicit operator string(DynamicXml xml)
        {
            return xml._xml.Value;
        }
    }
}
