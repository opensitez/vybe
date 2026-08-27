// vybe-test: csharp/csharp_with_expression_records/with_record_method_after
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var d=(new Counter(1)) with{N=5}
;
__P((d.Next()).ToString());
__Check("6");

record Counter(int N){public int Next()=>N+1;}

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
