// vybe-test: csharp/csharp_collection_expressions/collection_expression_modifying_copy_not_source
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

int[] src = [1, 2];
int[] copy = [..src];
copy[0] = 9;
__P((src[0]).ToString());
__P((copy[0]).ToString());
__Check("1\n9");

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
