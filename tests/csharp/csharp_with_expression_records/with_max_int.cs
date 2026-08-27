// vybe-test: csharp/csharp_with_expression_records/with_max_int
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var m=(new V(1)) with{N=int.MaxValue}
;
__P((m.N==int.MaxValue).ToString());
__Check("True");

record V(int N);

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
