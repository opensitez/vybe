// vybe-test: csharp/csharp_records_advanced/record_hash_code_matches_for_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var left = new Point(5, 7);
var right = new Point(5, 7);
__P((left.GetHashCode() == right.GetHashCode()).ToString());
__Check("True");

record Point(int X, int Y);

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
