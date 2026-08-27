// vybe-test: csharp/csharp_xml_linq/xdocument_root_element_accessible
// origin: languages/csharp/tests/csharp/test_csharp_xml_linq.rs

using static __Harness;

var doc=System.Xml.Linq.XDocument.Parse("<root><child>v</child></root>");
__P((doc.Root.Name.LocalName).ToString());
__Check("root");

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
