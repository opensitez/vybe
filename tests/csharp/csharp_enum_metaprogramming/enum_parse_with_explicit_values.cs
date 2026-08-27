// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_with_explicit_values
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var v=(Http)System.Enum.Parse(typeof(Http),"NotFound");
__P(((int)v).ToString());
__Check("404");

enum Http{Ok=200,NotFound=404}

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
