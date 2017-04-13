using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ConsoleApplication1
{
    class WeatherTxt
    {
        private FileStream BodyTxt;
        private string BodyStr = "";
        private string appId = File.ReadAllText(@"appId.txt");
        private string url = "";
        private string date = "";

        public WeatherTxt() { }
        public WeatherTxt(FileStream _BodyTxt, string _BodyStr)
        {
            BodyTxt = _BodyTxt;
            BodyStr = _BodyStr;
        }

        public void Clear()
        {
            BodyStr = "";
            url = "";
            date = "";

        }

        public void SetBody(string newBodyStr)
        {
            BodyStr = BodyStr + "\n" + newBodyStr;
        }

        public string GetDate()
        {
            return date;
        }

        public string getBody()
        {
            return BodyStr;
        }

        public FileStream getBodyTxt()
        {
            return BodyTxt;
        }

        public void BodyToTxt(string newBodyStr)
        {
            BodyTxt = new FileStream("weather.txt", FileMode.Append);
            StreamWriter writer = new StreamWriter(BodyTxt);
            BodyStr = newBodyStr;
            writer.WriteLine(BodyStr);
            Console.WriteLine(BodyStr);
            writer.Close();
        }

        public WeatherTxt GetWeather(string method, string city)
        {
            url = "http://api.openweathermap.org/data/2.5/" + method + "?q=" + city + "&mode=xml&units=metric" + appId;
            date = @"H:\Users\denis\\Documents\Visual Studio 2015\Projects\ConsoleApplication1\ConsoleApplication1\bin\\Debug\WeatherXml\" + Regex.Replace(System.DateTime.Now.ToLongTimeString(), "[:]", ".") + ".xml";
            Console.WriteLine(date);

            XDocument weather = new XDocument();
            weather = XDocument.Load(url);
            weather.Save(date);

            if (method == "weather")
            {
                foreach (XElement TemElement in weather.Element("current").Elements("temperature"))
                {
                    XAttribute valueAttribute = TemElement.Attribute("value");
                    SetBody(valueAttribute.Value);
                }
                BodyToTxt(getBody());
            }
            else if (method == "forecast")
            {
                foreach (XElement TimeElement in weather.Element("weatherdata").Element("forecast").Elements("time"))
                {
                    XAttribute valueAttribute = TimeElement.Element("temperature").Attribute("value");
                    SetBody(valueAttribute.Value);
                }
                BodyToTxt(getBody());
            }
            else if (method == "forecast/daily")
            {
                foreach (XElement TimeElement in weather.Element("weatherdata").Element("forecast").Elements("time"))
                {
                    XAttribute valueAttribute = TimeElement.Element("temperature").Attribute("day");
                    SetBody(valueAttribute.Value);
                }
                BodyToTxt(getBody());
            }

                WeatherTxt weatherTxt = new WeatherTxt(getBodyTxt(), getBody());

            return weatherTxt;
        }
    }
}
