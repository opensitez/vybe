// vybe-test: csharp/csharp_yield_iterators_core/yield_break_inside_nested_loop
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> GetNums() {
    for (int i = 0; i < 3; i++) {
        yield return i;
    }
}
string res = string.Join(",", GetNums());
__P(res);
__Check("0,1,2");
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
