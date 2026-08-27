// vybe-test: csharp/csharp_linq_complex/aggregate_with_seed_computes_running_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

using static __Harness;

var result=new[]{1,2,3,4,5}
.Aggregate(1L,(acc,n)=>acc*n);
__P((result).ToString());
__Check("120");

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
