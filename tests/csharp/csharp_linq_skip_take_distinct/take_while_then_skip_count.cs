// vybe-test: csharp/csharp_linq_skip_take_distinct/take_while_then_skip_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

int[] nums = new int[] { 1, 2, 3, 10, 11 };
int count = System.Linq.Enumerable.Count(System.Linq.Enumerable.Skip(System.Linq.Enumerable.TakeWhile(nums, x => x < 10), 1));
__P(count.ToString());
__Check("2");
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
