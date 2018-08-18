using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace computerman_rtg_reports
{
    public class Service
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }

        public static List<Service> getPricelist(string name)
        {
            string path = "Pricelists/"+name+".xml";
            XDocument doc = XDocument.Load(path);
            XElement root = doc.Element("Pricelist");
            
            List<Service> services = new List<Service>();

            foreach (XElement elem in root.Elements("Item"))
            {
                Service serv = new Service();
                serv.Id = Int32.Parse(elem.Attribute("Id").Value);
                serv.Code = elem.Attribute("Code").Value;
                serv.Name = elem.Attribute("Name").Value;
                serv.Price = decimal.Parse(elem.Attribute("Price").Value);

                services.Add(serv);
            }

            return services;
        }
    }
}