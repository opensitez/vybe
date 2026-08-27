// vybe-test: csharp/csharp_char_type_semantics/char_equality_compares_code_units_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

using static __Harness;

char left = 'Z';
char right = 'Z';
__P((left == right).ToString());
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
