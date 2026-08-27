// vybe-test: csharp/multi_value/tuple_return_and_deconstruct
// origin: languages/csharp/tests/csharp/test_multi_value.rs

using static __Harness;

__P("Valid_tuple_return_and_deconstruct");
__Check("Valid_tuple_return_and_deconstruct");
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
