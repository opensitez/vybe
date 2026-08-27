// vybe-test: csharp/csharp_xml_linq/xelement_construction_with_children_builds_tree
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

using static __Harness;

var xml=new System.Xml.Linq.XElement("Root",
    new System.Xml.Linq.XElement("Child","data"));
__P((xml.Element("Child").Value).ToString());
__Check("data");

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
