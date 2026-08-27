// vybe-test: csharp/csharp_nullable_ref_types/is_not_null_pattern_guards_against_null_dereference
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

using static __Harness;

string? s="hello";
if(s is not null) __P((s.Length).ToString());
__Check("5");

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
