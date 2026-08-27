// vybe-test: csharp/csharp_icomparable_sorting/comparer_create_builds_comparer_from_lambda
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

using static __Harness;

var cmp = System.Collections.Generic.Comparer<string>.Create(
    (a,b) => a.Length.CompareTo(b.Length));
var list = new System.Collections.Generic.List<string>{"cc","aaa","b"}
;
list.Sort(cmp);
__P((list[0]).ToString());
__Check("b");

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
