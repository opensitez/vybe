// vybe-test: csharp/csharp_char_type_semantics/new_string_from_char_array_reconstructs_text
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

using static __Harness;

char[] data = { 'h', 'i' }
;
__P((new string(data)).ToString());
__Check("hi");

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
