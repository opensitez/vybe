// vybe-test: csharp/csharp_expression_bodied/expression_bodied_property_getter
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

class Circle{public double R;public double Area=>System.Math.PI*R*R;}
__P((System.Math.Round(new Circle{R=0}.Area)).ToString());
__Check("0");
