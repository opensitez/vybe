// vybe-test: csharp/csharp_xml_linq/xelement_attribute_read_back
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

using static __Harness;

var el=new System.Xml.Linq.XElement("Node",
    new System.Xml.Linq.XAttribute("id","42"));
__P(((string)el.Attribute("id")).ToString());
__Check("42");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
