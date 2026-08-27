// vybe-test: csharp/csharp_array_apis/array_true_for_all_checks_entire_sequence
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var values = new[] { 2, 4, 6 }
;
__P((System.Array.TrueForAll(values, value => value % 2 == 0)).ToString());
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
