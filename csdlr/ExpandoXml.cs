using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace csdlr
{
    public static class ExpandoXml
    {
        /// <summary>
        /// Utworzenie ExpandoObject dla zadanego dokumentu XML
        /// </summary>
        /// <param name="document">Dokument XML, na podstawie którego zostanie utworzony ExpandoObject</param>
        /// <returns>ExpandoObject utworzony na podstawie zadanego dokumentu XML</returns>
        public static dynamic AsExpando(this XDocument document)
        {
            return CreateExpando(document.Root); //Utworzenie ExpandoObject dla elementu najwyższego w hierarchi w zadanym dokumencie XML
        }

        /// <summary>
        /// Utworzenie ExpandoObject dla zadanego elementu XML
        /// </summary>
        /// <param name="element">Element XML, na podstawie którego zostanie utworzony ExpandoObject</param>
        /// <returns>ExpandoObject utworzony na podstawie zadanego elementu XML</returns>
        private static dynamic CreateExpando(XElement element)
        {
            //Utworzenie instancji klasy ExpandoObject i zrzutowanie jej na interfejs IDictionary w którym kluczem jest string, a wartością object
            var result = new ExpandoObject() as IDictionary<string, object>;

            //Sprawdzenie czy jakikolwiek z elementów podrzędnych przekazanego elementu posiada podelementy - czy jakikolwiek element (węzeł) nie jest liściem
            if (element.Elements().Any(e => e.HasElements))
            {
                var list = new List<ExpandoObject>(); //Utworzenie listy obiektów ExpandoObject

                result.Add(element.Name.ToString(), list); //Dodanie do słownika nazwy przekazanego elementu i listy ExpandoObject

                foreach (var childElement in element.Elements()) //Przejście po każdym dziecku przekazanego elementu
                {
                    //Dodanie do listy obiektu ExpandoObject utworzonego na podstawie danego podelementu przekazanego elementu
                    list.Add(CreateExpando(childElement));
                }
            }
            else //Wszystkie elementy podrzędne przekazanego elementu są liśćmi - nie posiadają podelementów
            {
                foreach (var leafElement in element.Elements()) //Przejście po każdym podelemencie (każdy jest liściem) przekazanego elementu
                {
                    result.Add(leafElement.Name.ToString(), leafElement.Value); //Dodanie do słownika nazwy elementu i jego wartości
                }
            }

            return result; //Zwrócenie utworzonego ExpandoObject (słownika)
        }

        public static dynamic AsExpando2(this XDocument document)
        {
            return CreateExpando2(document.Root);
        }

        private static dynamic CreateExpando2(XElement element)
        {
            var result = new ExpandoObject() as IDictionary<string, object>;

            if (element.HasElements)
            {
                var list = new List<ExpandoObject>();

                result.Add(element.Name.LocalName, list);

                foreach (var childElement in element.Elements())
                {
                    list.Add(CreateExpando2(childElement));
                }
            }
            else
            {
                result.Add(element.Name.LocalName, element.Value);
            }

            return result;
        }

        public static dynamic ToExpando(this XDocument xDocument)
        {
            string jsonText = JsonConvert.SerializeXNode(xDocument);

            dynamic dyn = JsonConvert.DeserializeObject<ExpandoObject>(jsonText);

            return dyn;
        }

        public static dynamic Convert(this XDocument xDocument)
        {
            return Convert(xDocument.Root);
        }

        private static dynamic Convert(XElement parent)
        {
            dynamic output = new ExpandoObject();

            output.Name = parent.Name.LocalName;
            output.Value = parent.Value;

            output.HasAttributes = parent.HasAttributes;

            if (parent.HasAttributes)
            {
                output.Attributes = new List<KeyValuePair<string, string>>();

                foreach (XAttribute attr in parent.Attributes())
                {
                    KeyValuePair<string, string> temp = new KeyValuePair<string, string>(attr.Name.LocalName, attr.Value);
                    output.Attributes.Add(temp);
                }
            }

            output.HasElements = parent.HasElements;

            if (parent.HasElements)
            {
                output.Elements = new List<dynamic>();

                foreach (XElement element in parent.Elements())
                {
                    dynamic temp = Convert(element);
                    output.Elements.Add(temp);
                }
            }

            return output;
        }

        public static dynamic Parse(this XDocument xDocument)
        {
            dynamic parent = new ExpandoObject();

            Parse(parent, xDocument.Root);

            return parent;
        }

        public static void Parse(dynamic parent, XElement node)
        {
            if (node.HasElements)
            {
                if (node.Elements(node.Elements().First().Name.LocalName).Count() > 1)
                {
                    //list
                    var item = new ExpandoObject();
                    var list = new List<dynamic>();
                    foreach (var element in node.Elements())
                    {
                        Parse(list, element);
                    }

                    AddProperty(item, node.Elements().First().Name.LocalName, list);
                    AddProperty(parent, node.Name.ToString(), item);
                }
                else
                {
                    var item = new ExpandoObject();

                    foreach (var attribute in node.Attributes())
                    {
                        AddProperty(item, attribute.Name.ToString(), attribute.Value.Trim());
                    }

                    //element
                    foreach (var element in node.Elements())
                    {
                        Parse(item, element);
                    }

                    AddProperty(parent, node.Name.ToString(), item);
                }
            }
            else
            {
                AddProperty(parent, node.Name.ToString(), node.Value.Trim());
            }
        }

        private static void AddProperty(dynamic parent, string name, object value)
        {
            if (parent is List<dynamic>)
            {
                (parent as List<dynamic>).Add(value);
            }
            else
            {
                (parent as IDictionary<string, object>)[name] = value;
            }
        }
    }
}
