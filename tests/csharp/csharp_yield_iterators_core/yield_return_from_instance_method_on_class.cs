// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_instance_method_on_class
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

__P((new Counter().Range(4).Sum()).ToString());
__Check("6");

class Counter{public System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}}

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
