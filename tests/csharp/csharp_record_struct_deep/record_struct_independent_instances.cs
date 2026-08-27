// vybe-test: csharp/csharp_record_struct_deep/record_struct_independent_instances
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var a=new V(1);
var b=new V(2);
var c=a with{N=5}
;
__P((b.N).ToString());
__P((c.N).ToString());
__Check("2\n5");

record struct V(int N);

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
