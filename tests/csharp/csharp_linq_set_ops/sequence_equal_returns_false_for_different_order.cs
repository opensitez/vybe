// vybe-test: csharp/csharp_linq_set_ops/sequence_equal_returns_false_for_different_order
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

using static __Harness;

__P((new[]{1,2,3}.SequenceEqual(new[]{3,2,1})).ToString());
__Check("False");

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
