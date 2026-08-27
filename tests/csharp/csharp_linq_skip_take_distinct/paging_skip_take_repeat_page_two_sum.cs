// vybe-test: csharp/csharp_linq_skip_take_distinct/paging_skip_take_repeat_page_two_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

int[] nums = new int[] { 1, 2, 3, 4, 5, 6 };
int sum = System.Linq.Enumerable.Sum(System.Linq.Enumerable.Take(System.Linq.Enumerable.Skip(nums, 2), 2));
__P(sum.ToString());
__Check("7");
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
