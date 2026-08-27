// vybe-test: csharp/csharp_enum_operations/enum_cast_to_underlying_int_type
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

using static __Harness;

__P(((int)Priority.High).ToString());
__Check("3");

enum Priority{Low=1,Medium=2,High=3}

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
