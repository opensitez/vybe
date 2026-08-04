// vybe-test: csharp/csharp_struct_features/struct_equality_via_overridden_equals
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

struct Color {
    public int R,G,B;
    public override bool Equals(object o) => o is Color c && c.R==R && c.G==G && c.B==B;
    public override int GetHashCode() => System.HashCode.Combine(R,G,B);
}
var x = new Color{R=1,G=2,B=3};
var y = new Color{R=1,G=2,B=3};
__P((x.Equals(y)).ToString());
__Check("True");
