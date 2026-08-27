// vybe-test: csharp/csharp_enum_guid_version/version_compare_to_reports_negative_for_smaller_version
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var left = new System.Version(1, 2);
var right = new System.Version(1, 3);
__P((left.CompareTo(right)).ToString());
__Check("-1");

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
