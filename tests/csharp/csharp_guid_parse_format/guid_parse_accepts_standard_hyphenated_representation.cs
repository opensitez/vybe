// vybe-test: csharp/csharp_guid_parse_format/guid_parse_accepts_standard_hyphenated_representation
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

using static __Harness;

var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
__P((id.ToString().StartsWith("11111111")).ToString());
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
