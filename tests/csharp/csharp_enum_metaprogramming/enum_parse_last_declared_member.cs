// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_last_declared_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var v=(Level)System.Enum.Parse(typeof(Level),"High");
__P((v).ToString());
__Check("High");

enum Level{Low,Mid,High}

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
