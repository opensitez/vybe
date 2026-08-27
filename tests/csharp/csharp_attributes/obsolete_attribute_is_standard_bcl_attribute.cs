// vybe-test: csharp/csharp_attributes/obsolete_attribute_is_standard_bcl_attribute
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

using static __Harness;

var mi=typeof(Old).GetMethod("OldMethod");
bool hasObs=mi.GetCustomAttributes(typeof(System.ObsoleteAttribute),false).Length>0;
__P((hasObs).ToString());
__Check("True");

class Old{
    [System.Obsolete("use NewMethod")]
    public void OldMethod(){}
}

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
