// vybe-test: csharp/csharp_xml_linq/xdocument_root_element_accessible
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var doc=System.Xml.Linq.XDocument.Parse("<root><child>v</child></root>");
__Check((doc.Root.Name.LocalName).ToString(), "root");
