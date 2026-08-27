// vybe-test: csharp/csharp_yield_iterators_core/yield_return_filter_with_continue_skip
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Evens(int max){for(int i=0;i<=max;i++){if(i%2!=0)continue;yield return i;}}
__P((string.Join(",",Evens(6))).ToString());
__Check("0,2,4,6");

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
