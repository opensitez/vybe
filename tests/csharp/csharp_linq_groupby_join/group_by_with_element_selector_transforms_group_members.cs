// vybe-test: csharp/csharp_linq_groupby_join/group_by_with_element_selector_transforms_group_members
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

using static __Harness;

var nums = new[] { 1, 2, 3, 4 }
;
var groups = nums.GroupBy(n => n % 2 == 0 ? "even" : "odd",
                          n => n * 10);
int evenSum = 0;
foreach (var g in groups)
    if (g.Key == "even") foreach (var v in g) evenSum += v;
__P((evenSum).ToString());
__Check("60");

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
