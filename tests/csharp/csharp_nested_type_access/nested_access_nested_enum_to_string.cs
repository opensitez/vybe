// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_to_string
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Mode().Label()).ToString());
__Check("Beta");

class Mode{public enum Kind{Alpha,Beta} public string Label()=>Kind.Beta.ToString();}

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
