// vybe-test: csharp/csharp_xml_linq/xelement_descendants_query_returns_matching_nodes
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var doc=System.Xml.Linq.XDocument.Parse("<root><a>1</a><a>2</a></root>");
int count=doc.Root.Elements("a").Count();
__P((count).ToString());
__Check("2");
