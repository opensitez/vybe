// vybe-test: csharp/csharp_string_parsing/guid_try_parse_recognises_standard_format
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

using static __Harness;

__P((System.Guid.TryParse("550e8400-e29b-41d4-a716-446655440000",out _)).ToString());
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
