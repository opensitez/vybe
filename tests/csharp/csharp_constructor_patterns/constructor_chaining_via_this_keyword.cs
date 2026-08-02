// vybe-test: csharp/csharp_constructor_patterns/constructor_chaining_via_this_keyword
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box{public int W,H,D;
    public Box(int w,int h,int d){W=w;H=h;D=d;}
    public Box(int side):this(side,side,side){}
}
var cube=new Box(3);
__Check((cube.W).ToString(), "3"); __Check((cube.H).ToString(), "3"); __Check((cube.D).ToString(), "3");
