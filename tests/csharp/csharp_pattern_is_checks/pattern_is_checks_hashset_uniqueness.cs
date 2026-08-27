// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

using static __Harness;

// pattern_is_checks
var set = new System.Collections.Generic.HashSet<int>();
set.Add(41);
set.Add(41);
__P((set.Count == 1).ToString());
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
