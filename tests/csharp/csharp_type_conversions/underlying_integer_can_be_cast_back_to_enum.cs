// vybe-test: csharp/csharp_type_conversions/underlying_integer_can_be_cast_back_to_enum
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;

var mode = (Mode)5;
__P((mode).ToString());
__Check("On");

enum Mode { Off = 0, On = 5 }

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
