// vybe-test: csharp/csharp_control_flow/foreach_on_dictionary_visits_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

var map = new System.Collections.Generic.Dictionary<string, int> { ["x"] = 1 }
;
int total = 0;
foreach (var pair in map) total += pair.Value;
__P((total).ToString());
__Check("1");

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
