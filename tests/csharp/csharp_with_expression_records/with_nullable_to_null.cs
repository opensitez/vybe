// vybe-test: csharp/csharp_with_expression_records/with_nullable_to_null
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var z=(new Maybe(5)) with{N=null}
;
__P((z.N.HasValue).ToString());
__Check("False");

record Maybe(int? N);

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
