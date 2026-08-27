// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_explicit_ienumerable_interface
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

__P((new Nums().Sum()).ToString());
__Check("6");

class Nums:System.Collections.Generic.IEnumerable<int>{public System.Collections.Generic.IEnumerator<int> GetEnumerator(){yield return 2;yield return 4;}System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();}

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
