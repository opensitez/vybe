// vybe-test: csharp/csharp_xml_linq/xelement_attribute_read_back
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var el=new System.Xml.Linq.XElement("Node",
    new System.Xml.Linq.XAttribute("id","42"));
__Check(((string)el.Attribute("id")).ToString(), "42");
