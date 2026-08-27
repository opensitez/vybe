// vybe-test: csharp/csharp_closures/nested_closure_captures_from_outer_scope
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

using static __Harness;

System.Func<int,System.Func<int>> makeAdder = x => () => x + 1;
var add1 = makeAdder(5);
__P((add1()).ToString());
__Check("6");

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
