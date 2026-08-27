// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_parameterized_start_index
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> From(int start,int count){for(int i=0;i<count;i++)yield return start+i;}
__P((string.Join(",",From(5,3))).ToString());
__Check("5,6,7");

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
