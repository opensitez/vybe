// vybe-test: csharp/csharp_record_struct/record_struct_equality_compares_by_value
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

using static __Harness;

var c1=new Color(255,0,0);
var c2=new Color(255,0,0);
__P((c1==c2).ToString());
__Check("True");

record struct Color(int R,int G,int B);

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
