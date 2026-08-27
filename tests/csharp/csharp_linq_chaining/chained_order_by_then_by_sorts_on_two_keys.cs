// vybe-test: csharp/csharp_linq_chaining/chained_order_by_then_by_sorts_on_two_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

using static __Harness;

var data=new[]{(A:"b",B:2),(A:"a",B:3),(A:"a",B:1)}
;
var result=data.OrderBy(x=>x.A).ThenBy(x=>x.B);
foreach(var(a,b) in result) __P(($"{a}{b}").ToString());
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
