// vybe-test: csharp/advanced/array_foreach_sum
// origin: languages/csharp/tests/csharp/test_advanced.rs

using static __Harness;

var nums = new int[] { 10, 20, 30, 40 }
;
var total = 0;
foreach (var n in nums) { total = total + n; }
__P((total).ToString());
__Check("100");

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
