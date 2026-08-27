// vybe-test: csharp/csharp_iterators/yield_in_loop_produces_computed_sequence
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Range(int n) {
    for(int i=0; i<n; i++) yield return i;
}
int sum=0;
foreach(var x in Range(5)) sum+=x;
__P((sum).ToString());
__Check("10");

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
