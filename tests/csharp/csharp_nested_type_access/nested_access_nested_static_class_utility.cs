// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_static_class_utility
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((Text.Merge()).ToString());
__Check("ab");

class Text{public static class Util{public static string Join(string a,string b)=>a+b;} public static string Merge()=>Util.Join("a","b");}

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
