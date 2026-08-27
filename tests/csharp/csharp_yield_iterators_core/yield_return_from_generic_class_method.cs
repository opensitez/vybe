// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_generic_class_method
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

__P((new Bag<int>().Single(8).First()).ToString());
__Check("8");

class Bag<T>{public System.Collections.Generic.IEnumerable<T> Single(T v){yield return v;}}

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
