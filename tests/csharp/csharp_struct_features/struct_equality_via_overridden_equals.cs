// vybe-test: csharp/csharp_struct_features/struct_equality_via_overridden_equals
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((x.Equals(y)).ToString(), "True");
