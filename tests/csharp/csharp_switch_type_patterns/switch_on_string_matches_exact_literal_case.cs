// vybe-test: csharp/csharp_switch_type_patterns/switch_on_string_matches_exact_literal_case
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

string Pick(string key) {
    switch (key) {
        case "go": return "G";
        case "stop": return "S";
        default: return "?";
    }
}
__P((Pick("go")).ToString());
__Check("G");

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
