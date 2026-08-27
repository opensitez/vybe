// vybe-test: csharp/csharp_nullable_ref_types/nullable_reference_getvalueorlength_via_null_coalescing
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

using static __Harness;

string? s=null;
int len=s?.Length??-1;
__P((len).ToString());
__Check("-1");

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
