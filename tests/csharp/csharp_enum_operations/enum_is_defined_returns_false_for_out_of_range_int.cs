// vybe-test: csharp/csharp_enum_operations/enum_is_defined_returns_false_for_out_of_range_int
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

using static __Harness;

__P((System.Enum.IsDefined(typeof(Level), 99)).ToString());
__Check("False");

enum Level{Low=0,Mid=1,High=2}

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
