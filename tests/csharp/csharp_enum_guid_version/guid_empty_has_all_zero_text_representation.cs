// vybe-test: csharp/csharp_enum_guid_version/guid_empty_has_all_zero_text_representation
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

__P((System.Guid.Empty.ToString()).ToString());
__Check("00000000-0000-0000-0000-000000000000");

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
