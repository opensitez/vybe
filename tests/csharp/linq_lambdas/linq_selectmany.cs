// vybe-test: csharp/linq_lambdas/linq_selectmany
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var lists = new List<List<int>> {
    new List<int> { 1, 2 },
    new List<int> { 3, 4 },
    new List<int> { 5 }
}
;
var flat = lists.SelectMany(l => l).ToList();
foreach (var x in flat) __P((x).ToString());
__Check("1\n2\n3\n4\n5");

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
