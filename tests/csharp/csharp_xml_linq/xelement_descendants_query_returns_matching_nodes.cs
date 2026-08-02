// vybe-test: csharp/csharp_xml_linq/xelement_descendants_query_returns_matching_nodes
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var doc=System.Xml.Linq.XDocument.Parse("<root><a>1</a><a>2</a></root>");
int count=doc.Root.Elements("a").Count();
__Check((count).ToString(), "2");
