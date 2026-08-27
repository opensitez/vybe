// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_chain
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var a=new Box(1,1);
var b=a with{W=2}
;
var c=b with{H=3}
;
__P((a.W).ToString());
__P((c.W).ToString());
__P((c.H).ToString());
__Check("1\n2\n3");

record struct Box(int W,int H);

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
