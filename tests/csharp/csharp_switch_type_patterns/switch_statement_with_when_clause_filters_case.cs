// vybe-test: csharp/csharp_switch_type_patterns/switch_statement_with_when_clause_filters_case
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

int n = 8;
string size = n switch {
    < 0 => "neg",
    >= 0 and < 10 => "small",
    _ => "big"
}
;
__P((size).ToString());
__Check("small");

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
