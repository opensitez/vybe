// vybe-test: csharp/csharp_strings_ext/string_intern_returns_reference_equal_for_same_literal_sequence
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

using static __Harness;

string a = string.Intern("shared");
string b = string.Intern("shared");
__P((object.ReferenceEquals(a, b)).ToString());
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
