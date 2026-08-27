// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

using static __Harness;

// exception_type_checks
var set = new System.Collections.Generic.HashSet<int>();
set.Add(53);
set.Add(53);
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
