// vybe-test: csharp/csharp_yield_iterators_core/nested_three_level_iterator_chain
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> A(){yield return 1;}
System.Collections.Generic.IEnumerable<int> B(){foreach(var x in A())yield return x+10;}
System.Collections.Generic.IEnumerable<int> C(){foreach(var x in B())yield return x+100;}
__P((string.Join(",",C())).ToString());
__Check("111");

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
