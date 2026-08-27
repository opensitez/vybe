// vybe-test: csharp/csharp_with_expression_records/with_nominal_chain
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var e=((new C{A=1,B=2}) with{A=3}) with{B=4}
;
__P((e.A).ToString());
__P((e.B).ToString());
__Check("3\n4");

record C{public int A{get;init;} public int B{get;init;}}

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
