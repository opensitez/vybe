// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_readonly
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var s=new Size(2,3);
var t=s with{H=8}
;
__P((s.H).ToString());
__P((t.H).ToString());
__Check("3\n8");

readonly record struct Size(int W,int H);

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
