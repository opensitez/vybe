// vybe-test: csharp/csharp_enum_guid_version/enum_has_flag_detects_enabled_bit_flag
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var value = Permission.Read | Permission.Write;
__P((value.HasFlag(Permission.Write)).ToString());
__Check("True");

[System.Flags] enum Permission { Read = 1, Write = 2, Execute = 4 }

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
