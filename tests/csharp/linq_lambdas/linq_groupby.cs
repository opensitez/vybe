// vybe-test: csharp/linq_lambdas/linq_groupby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var words = new List<string> { "apple", "ant", "banana", "avocado", "bat" }
;
var groups = words.GroupBy(w => w[0].ToString()).ToList();
foreach (var g in groups) {
    __P((g.Key + ": " + g.Count()).ToString());
}
__Check("a: 3\nb: 2");

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
