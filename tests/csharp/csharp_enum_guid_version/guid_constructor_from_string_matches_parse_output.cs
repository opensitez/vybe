// vybe-test: csharp/csharp_enum_guid_version/guid_constructor_from_string_matches_parse_output
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var text = "11111111-2222-3333-4444-555555555555";
__P((new System.Guid(text).ToString()).ToString());
__Check("11111111-2222-3333-4444-555555555555");

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
