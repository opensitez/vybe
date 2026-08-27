// vybe-test: csharp/csharp_expression_bodied/expression_bodied_constructor
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

var p=new Point(3,4);
__P((p.X).ToString());
__P((p.Y).ToString());
__Check("3\n4");

class Point{public int X,Y; public Point(int x,int y)=>(X,Y)=(x,y);}

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
