// vybe-test: csharp/csharp_pattern_matching/relational_pattern_compares_numeric_range
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

using static __Harness;

int score = 85;
string grade = score switch { >= 90 => "A", >= 80 => "B", >= 70 => "C", _ => "F" }
;
__P((grade).ToString());
__Check("B");

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
