// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

using static __Harness;

// exception_type_checks
var map = new System.Collections.Generic.Dictionary<int, int>();
map[53] = 54;
__P((map.ContainsKey(53) && map[53] == 54).ToString());
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
