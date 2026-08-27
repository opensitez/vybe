// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_from_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

var p=new Map().Origin();
__P((p.X).ToString());
__Check("0");

class Map{public struct Point{public int X; public int Y;} public Point Origin()=>new Point{X=0,Y=0};}

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
