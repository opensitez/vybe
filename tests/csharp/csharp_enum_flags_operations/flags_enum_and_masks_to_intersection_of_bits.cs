// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_and_masks_to_intersection_of_bits
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

using static __Harness;

var combined = Perm.A | Perm.B | Perm.C;
var masked = combined & Perm.B;
__P(((int)masked).ToString());
__Check("2");

[System.Flags]
enum Perm { A = 1, B = 2, C = 4 }

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
