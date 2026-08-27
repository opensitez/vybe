// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_different_element_same_length
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

__P((new[]{1,2,3}.SequenceEqual(new[]{1,2,4})).ToString());
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
