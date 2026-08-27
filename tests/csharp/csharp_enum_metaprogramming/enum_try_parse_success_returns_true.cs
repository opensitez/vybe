// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_success_returns_true
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var ok=System.Enum.TryParse<Day>("Tue",out var d);
__P((ok).ToString());
__P((d).ToString());
__Check("True\nTue");

enum Day{Mon,Tue,Wed}

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
