// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_local_state_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Running(){int s=0; for(int i=1;i<=3;i++){s+=i;yield return s;}}
__P((string.Join(",",Running())).ToString());
__Check("1,3,6");

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
