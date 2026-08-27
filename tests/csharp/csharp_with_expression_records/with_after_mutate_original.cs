// vybe-test: csharp/csharp_with_expression_records/with_after_mutate_original
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var a=new Box{V=1}
;
a.V=3;
var b=a with{V=4}
;
__P((b.V).ToString());
__Check("4");

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
