// vybe-test: csharp/csharp_enum_operations/flags_enum_has_flag_detects_combined_bit
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

using static __Harness;

var p = Perm.Read | Perm.Write;
__P((p.HasFlag(Perm.Read)).ToString());
__P((p.HasFlag(Perm.Execute)).ToString());
__Check("True\nFalse");

[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}

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
