// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_record_style_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Orders().Make(4).Total()).ToString());
__Check("8");

class Orders{public class Line{public int Qty; public int Total()=>Qty*2;} public Line Make(int q)=>new Line{Qty=q};}

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
