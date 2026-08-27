// vybe-test: csharp/csharp_with_expression_records/with_mutable_separate_instance
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var b=(new Box{V=1}) with{V=2}
;
__P((b.V).ToString());
__Check("2");

record Box{public int V{get;set;}}

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
