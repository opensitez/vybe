// vybe-test: csharp/csharp_struct_features/struct_method_operates_on_own_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

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

struct Vector { public double X,Y; public double Length() => System.Math.Sqrt(X*X+Y*Y); }
var v = new Vector { X=3, Y=4 };
__P((v.Length()).ToString());
__Check("5");
