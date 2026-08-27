// vybe-test: csharp/linq_runtime/linq_where_even
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

using static __Harness;

int[] nums = new int[] { 1, 2, 3, 4, 5 };
var evens = System.Linq.Enumerable.ToList(System.Linq.Enumerable.Where(nums, x => x % 2 == 0));
__P(evens.Count.ToString());
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
