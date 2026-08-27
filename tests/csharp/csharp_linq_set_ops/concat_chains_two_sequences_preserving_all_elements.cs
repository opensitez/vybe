// vybe-test: csharp/csharp_linq_set_ops/concat_chains_two_sequences_preserving_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

using static __Harness;

var result = new[]{1,2}
.Concat(new[]{3,4});
__P((result.Count()).ToString());
__Check("4");

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
