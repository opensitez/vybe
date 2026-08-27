// vybe-test: csharp/csharp_string_parsing/enum_try_parse_succeeds_on_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

using static __Harness;

__P((System.Enum.TryParse<Color>("Green",out var c)).ToString());
__P((c).ToString());
__Check("True\nGreen");

enum Color{Red,Green,Blue}

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
