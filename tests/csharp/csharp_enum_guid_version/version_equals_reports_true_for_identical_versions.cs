// vybe-test: csharp/csharp_enum_guid_version/version_equals_reports_true_for_identical_versions
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

__P((new System.Version(3, 5).Equals(new System.Version(3, 5))).ToString());
__Check("True");

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
