// vybe-test: csharp/csharp_collections_initialise/collection_expression_empty_array_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

using static __Harness;

int[] empty=[];
__P((empty.Length).ToString());
__Check("0");

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
