// vybe-test: csharp/csharp_struct_advanced/struct_iequatable_implementation_compares_by_value
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

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

struct Color:System.IEquatable<Color>{
    public int R,G,B;
    public bool Equals(Color o)=>R==o.R&&G==o.G&&B==o.B;
    public override bool Equals(object o)=>o is Color c&&Equals(c);
    public override int GetHashCode()=>System.HashCode.Combine(R,G,B);
}
var red1=new Color{R=255,G=0,B=0};
var red2=new Color{R=255,G=0,B=0};
__P((red1.Equals(red2)).ToString());
__Check("True");
