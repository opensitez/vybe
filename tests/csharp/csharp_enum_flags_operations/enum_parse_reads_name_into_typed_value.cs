// vybe-test: csharp/csharp_enum_flags_operations/enum_parse_reads_name_into_typed_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

using static __Harness;

var value = System.Enum.Parse(typeof(Color), "Green");
__P((value).ToString());
__Check("Green");

enum Color { Red, Green, Blue }

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
