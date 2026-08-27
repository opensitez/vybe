// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_length_counts_all_characters_including_escapes_as_literals
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

using static __Harness;

string s = @"\n";
__P(s.Length.ToString());
__Check("2");
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
