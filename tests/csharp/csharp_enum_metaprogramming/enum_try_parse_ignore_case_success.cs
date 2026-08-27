// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_ignore_case_success
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var ok=System.Enum.TryParse<Mode>("beta",true,out var m);
__P((ok).ToString());
__P((m).ToString());
__Check("True\nBeta");

enum Mode{Alpha,Beta}

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
