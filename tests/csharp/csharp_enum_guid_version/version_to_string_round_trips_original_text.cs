// vybe-test: csharp/csharp_enum_guid_version/version_to_string_round_trips_original_text
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var version = new System.Version(1, 2, 3, 4);
__P((version.ToString()).ToString());
__Check("1.2.3.4");

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
