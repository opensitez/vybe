// vybe-test: csharp/csharp_xml_linq/xelement_name_and_value_readable
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var el=new System.Xml.Linq.XElement("Item","hello");
__Check((el.Name.LocalName).ToString(), "Item"); __Check(((string)el).ToString(), "hello");
