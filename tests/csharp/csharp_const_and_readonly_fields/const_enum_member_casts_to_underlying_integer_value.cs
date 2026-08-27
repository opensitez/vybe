// vybe-test: csharp/csharp_const_and_readonly_fields/const_enum_member_casts_to_underlying_integer_value
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

using static __Harness;

const Code status = Code.Ok;
__P(((int)status).ToString());
__Check("0");

enum Code { Ok = 0, Err = 1 }

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
