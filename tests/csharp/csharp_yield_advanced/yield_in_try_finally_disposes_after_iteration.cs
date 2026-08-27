// vybe-test: csharp/csharp_yield_advanced/yield_in_try_finally_disposes_after_iteration
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

using static __Harness;

bool cleaned=false;
System.Collections.Generic.IEnumerable<int> Gen(){
    try{ yield return 1; yield return 2; }
    finally{ cleaned=true; }
}
foreach(var _ in Gen()){}
__P((cleaned).ToString());
__Check("True");

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
