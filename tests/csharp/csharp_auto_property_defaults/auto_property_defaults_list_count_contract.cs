// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

using static __Harness;

// auto_property_defaults
var values = new System.Collections.Generic.List<int> { 65, 66, 65 }
;
__P((values.Count == 3).ToString());
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
