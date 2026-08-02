// vybe-test: csharp/csharp_constructor_patterns/static_factory_method_creates_instance
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Color{
    public int R,G,B;
    public static Color FromGray(int v)=>new Color{R=v,G=v,B=v};
}
var gray=Color.FromGray(128);
__Check((gray.R==gray.G&&gray.G==gray.B).ToString(), "True");
