// vybe-test: csharp/csharp_xml_linq/xelement_construction_with_children_builds_tree
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

var xml=new System.Xml.Linq.XElement("Root",
    new System.Xml.Linq.XElement("Child","data"));
__P((xml.Element("Child").Value).ToString());
__Check("data");
