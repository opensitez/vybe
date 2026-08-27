// vybe-test: csharp/csharp_functional_patterns/map_then_filter_then_reduce_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

var result=new[]{1,2,3,4,5}
.Select(x=>x*x)
    .Where(x=>x>5)
    .Sum();
__P((result).ToString());
__Check("50");

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
