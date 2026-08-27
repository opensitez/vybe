// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

using static __Harness;

// nullable_value_operators
int? maybe = null;
int fallback = maybe ?? 57;
__P((fallback == 57).ToString());
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
