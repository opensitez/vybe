// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_yields_from_inner_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Inner(){yield return 1;yield return 2;}
System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in Inner())yield return x;foreach(var x in Inner())yield return x;}
__P((string.Join(",",Outer())).ToString());
__Check("1,2,1,2");

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
