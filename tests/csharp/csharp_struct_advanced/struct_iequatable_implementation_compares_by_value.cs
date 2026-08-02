// vybe-test: csharp/csharp_struct_advanced/struct_iequatable_implementation_compares_by_value
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((red1.Equals(red2)).ToString(), "True");
