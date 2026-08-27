// vybe-test: csharp/modern_features/index_from_end
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

int[] nums = { 10, 20, 30, 40, 50 }
;
__P((nums[^1]).ToString());
__P((nums[^2]).ToString());
__Check("50\n40");

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
