// vybe-test: csharp/csharp_ienumerable_custom/manual_enumerator_move_next_and_current
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

using static __Harness;

var list = new System.Collections.Generic.List<int> { 10, 20 };
var enumerator = list.GetEnumerator();
__P(enumerator.MoveNext().ToString());
__P(enumerator.Current.ToString());
__Check("True\n10");
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
