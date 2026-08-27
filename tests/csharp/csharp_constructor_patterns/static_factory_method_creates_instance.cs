// vybe-test: csharp/csharp_constructor_patterns/static_factory_method_creates_instance
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

using static __Harness;

var gray=Color.FromGray(128);
__P((gray.R==gray.G&&gray.G==gray.B).ToString());
__Check("True");

class Color{
    public int R,G,B;
    public static Color FromGray(int v)=>new Color{R=v,G=v,B=v};
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
