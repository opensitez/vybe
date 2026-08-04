// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_get_only_from_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

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

class Circle { public double R = 2.0; public double Area => System.Math.PI * R * R; }
__P((System.Math.Round(new Circle().Area, 2)).ToString());
__Check("12.57");
