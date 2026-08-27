// vybe-test: csharp/csharp_generics_constraints/generic_method_can_return_tuple_of_type_argument_and_count
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

(T, int) Pair<T>(T value) { return (value, 1); }
var result = Pair("x");
__P((result.Item1 + result.Item2).ToString());
__Check("x1");

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
