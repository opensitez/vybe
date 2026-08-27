// vybe-test: csharp/csharp_linq_projections/take_while_stops_at_first_failing_predicate
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var result = new[]{1,3,5,4,7}
.TakeWhile(x => x%2!=0);
foreach(var n in result) __P((n).ToString());
__Check("1\n3\n5");

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
