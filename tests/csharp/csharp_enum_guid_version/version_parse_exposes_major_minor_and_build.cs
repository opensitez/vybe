// vybe-test: csharp/csharp_enum_guid_version/version_parse_exposes_major_minor_and_build
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var version = System.Version.Parse("2.4.6");
__P((version.Major).ToString());
__P((version.Minor).ToString());
__P((version.Build).ToString());
__Check("2\n4\n6");

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
