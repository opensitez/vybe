// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_then_cast_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var p=(Round)System.Enum.Parse(typeof(Round),"B");
__P(((int)p).ToString());
__Check("22");

enum Round{A=11,B=22}

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
