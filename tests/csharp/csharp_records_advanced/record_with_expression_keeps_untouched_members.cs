// vybe-test: csharp/csharp_records_advanced/record_with_expression_keeps_untouched_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var item = new Item("pen", 2);
var changed = item with { Count = 5 }
;
__P((changed.Name).ToString());
__Check("pen");

record Item(string Name, int Count);

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
