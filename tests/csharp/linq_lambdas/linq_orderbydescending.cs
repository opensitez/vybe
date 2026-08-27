// vybe-test: csharp/linq_lambdas/linq_orderbydescending
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var nums = new List<int> { 3, 1, 4, 1, 5 }
;
var sorted = nums.OrderByDescending(x => x).ToList();
foreach (var x in sorted) __P((x).ToString());
__Check("5\n4\n3\n1\n1");

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
