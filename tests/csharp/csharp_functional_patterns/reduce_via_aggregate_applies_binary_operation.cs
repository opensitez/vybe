// vybe-test: csharp/csharp_functional_patterns/reduce_via_aggregate_applies_binary_operation
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

var product=new[]{1,2,3,4,5}
.Aggregate((acc,x)=>acc*x);
__P((product).ToString());
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
