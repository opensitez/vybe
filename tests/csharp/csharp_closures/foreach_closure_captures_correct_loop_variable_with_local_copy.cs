// vybe-test: csharp/csharp_closures/foreach_closure_captures_correct_loop_variable_with_local_copy
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

using static __Harness;

var actions = new System.Collections.Generic.List<System.Func<int>>();
foreach(var v in new[]{10,20,30}) {
    var copy = v;
    actions.Add(() => copy);
}
foreach(var a in actions) __P((a()).ToString());
__Check("10\n20\n30");

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
