// vybe-test: csharp/csharp_constructor_patterns/constructor_chaining_via_this_keyword
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Box{public int W,H,D;
    public Box(int w,int h,int d){W=w;H=h;D=d;}
    public Box(int side):this(side,side,side){}
}
var cube=new Box(3);
__P((cube.W).ToString()); __P((cube.H).ToString()); __P((cube.D).ToString());
__Check("3\n3\n3");
