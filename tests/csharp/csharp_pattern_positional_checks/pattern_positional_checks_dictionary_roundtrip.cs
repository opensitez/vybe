// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

using static __Harness;

// pattern_positional_checks
var map = new System.Collections.Generic.Dictionary<int, int>();
map[115] = 116;
__P((map.ContainsKey(115) && map[115] == 116).ToString());
__Check("True");

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
