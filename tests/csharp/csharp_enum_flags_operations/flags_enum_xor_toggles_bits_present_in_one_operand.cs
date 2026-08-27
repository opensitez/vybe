// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_xor_toggles_bits_present_in_one_operand
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

using static __Harness;

var value = (Perm.A | Perm.B) ^ Perm.A;
__P(((int)value).ToString());
__Check("2");

[System.Flags]
enum Perm { A = 1, B = 2 }

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
