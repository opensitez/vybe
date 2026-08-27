// vybe-test: csharp/csharp_linq_let_join/order_by_in_query_syntax_sorts_ascending
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

using static __Harness;

var q=from n in new[]{3,1,2}
orderby n select n;
foreach(var x in q) __P((x).ToString());
__Check("1\n2\n3");

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
