// vybe-test: csharp/csharp_with_expression_records/with_nullable_to_value
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var v=(new Maybe(null)) with{N=7}
;
__P((v.N).ToString());
__Check("7");

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
