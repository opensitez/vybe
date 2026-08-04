// vybe-test: csharp/csharp_class_features/memberwise_clone_creates_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

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

class Point:System.ICloneable{public int X,Y;public object Clone()=>MemberwiseClone();}
var a=new Point{X=1,Y=2};
var b=(Point)a.Clone();
b.X=99;
__P((a.X).ToString());
__Check("1");
