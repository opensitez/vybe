// vybe-test: csharp/csharp_xml_linq/xelement_construction_with_children_builds_tree
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var xml=new System.Xml.Linq.XElement("Root",
    new System.Xml.Linq.XElement("Child","data"));
__Check((xml.Element("Child").Value).ToString(), "data");
