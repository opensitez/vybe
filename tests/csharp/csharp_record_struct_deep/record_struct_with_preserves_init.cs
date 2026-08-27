// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_preserves_init
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var p=new Pair{A=1,B=2}
;
var q=p with{A=9}
;
__P((q.B).ToString());
__Check("2");

record struct Pair{public int A{get;init;} public int B{get;init;}}

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
