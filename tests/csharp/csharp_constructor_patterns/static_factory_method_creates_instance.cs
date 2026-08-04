// vybe-test: csharp/csharp_constructor_patterns/static_factory_method_creates_instance
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

class Color{
    public int R,G,B;
    public static Color FromGray(int v)=>new Color{R=v,G=v,B=v};
}
var gray=Color.FromGray(128);
__P((gray.R==gray.G&&gray.G==gray.B).ToString());
__Check("True");
