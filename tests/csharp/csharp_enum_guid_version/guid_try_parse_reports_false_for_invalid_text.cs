// vybe-test: csharp/csharp_enum_guid_version/guid_try_parse_reports_false_for_invalid_text
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var ok = System.Guid.TryParse("bad-guid", out var value);
__P((ok).ToString());
__Check("False");

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
