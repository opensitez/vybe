// vybe-test: csharp/csharp_raw_string_literals/raw_string_custom_delimiter_single_quote
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

using static __Harness;

string raw = """
Content_raw_string_custom_delimiter_single_quote
""";
__P(raw.Trim());
__Check("Content_raw_string_custom_delimiter_single_quote");
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
