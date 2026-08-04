// vybe-test: csharp/csharp_expression_bodied/expression_bodied_constructor
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

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

class Point{public int X,Y; public Point(int x,int y)=>(X,Y)=(x,y);}
var p=new Point(3,4);
__P((p.X).ToString()); __P((p.Y).ToString());
__Check("3\n4");
