// vybe-test: csharp/csharp_enum_flags_operations/enum_switch_dispatches_on_underlying_constant_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

using static __Harness;

string Label(Mode mode) {
    switch (mode) {
        case Mode.Alpha: return "a";
        case Mode.Beta: return "b";
        default: return "?";
    }
}
__P((Label(Mode.Beta)).ToString());
__Check("b");

enum Mode { Alpha = 1, Beta = 2 }

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
