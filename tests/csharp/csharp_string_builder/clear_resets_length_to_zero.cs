// vybe-test: csharp/csharp_string_builder/clear_resets_length_to_zero
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

using static __Harness;

var sb = new System.Text.StringBuilder("data");
sb.Clear();
__P((sb.Length).ToString());
__Check("0");

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
