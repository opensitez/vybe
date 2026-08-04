// vybe-test: csharp/csharp_xml_linq/xelement_name_and_value_readable
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

var el=new System.Xml.Linq.XElement("Item","hello");
__P((el.Name.LocalName).ToString()); __P(((string)el).ToString());
__Check("Item\nhello");
