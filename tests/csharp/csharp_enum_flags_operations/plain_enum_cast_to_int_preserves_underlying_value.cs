// vybe-test: csharp/csharp_enum_flags_operations/plain_enum_cast_to_int_preserves_underlying_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

using static __Harness;

__P(((int)Level.Mid).ToString());
__Check("5");

enum Level { Low = 1, Mid = 5, High = 9 }

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
