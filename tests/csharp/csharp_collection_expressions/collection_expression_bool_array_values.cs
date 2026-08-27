// vybe-test: csharp/csharp_collection_expressions/collection_expression_bool_array_values
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

bool[] flags = [true, false, true];
__P((flags[0]).ToString());
__P((flags[1]).ToString());
__Check("True\nFalse");

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
