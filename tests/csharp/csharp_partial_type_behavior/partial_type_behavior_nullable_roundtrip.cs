// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

using static __Harness;

// partial_type_behavior
int? maybe = 70;
__P((maybe.HasValue && maybe.Value == 70).ToString());
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
