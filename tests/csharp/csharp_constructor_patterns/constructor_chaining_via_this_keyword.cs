// vybe-test: csharp/csharp_constructor_patterns/constructor_chaining_via_this_keyword
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

using static __Harness;

var cube=new Box(3);
__P((cube.W).ToString());
__P((cube.H).ToString());
__P((cube.D).ToString());
__Check("3\n3\n3");

class Box{public int W,H,D;
    public Box(int w,int h,int d){W=w;H=h;D=d;}
    public Box(int side):this(side,side,side){}
}

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
