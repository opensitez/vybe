// vybe-test: csharp/csharp_random_random/seeded_random_produces_deterministic_sequence
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

using static __Harness;

var r1=new System.Random(99);
var r2=new System.Random(99);
__P((r1.Next()==r2.Next()).ToString());
__Check("True");

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
