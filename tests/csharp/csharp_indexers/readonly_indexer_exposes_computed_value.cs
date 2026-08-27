// vybe-test: csharp/csharp_indexers/readonly_indexer_exposes_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

using static __Harness;

__P((new Odds()[4]).ToString());
__Check("9");

class Odds{public int this[int n]=>2*n+1;}

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
