// vybe-test: csharp/csharp_linq_groupby_join/order_by_then_by_applies_secondary_sort_on_equal_primary_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

using static __Harness;

var items = new[] { (Name:"b",Age:2),(Name:"a",Age:3),(Name:"a",Age:1) }
;
var sorted = items.OrderBy(x => x.Name).ThenBy(x => x.Age);
foreach (var x in sorted) __P(($"{x.Name}{x.Age}").ToString());
__Check("a1\na3\nb2");

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
