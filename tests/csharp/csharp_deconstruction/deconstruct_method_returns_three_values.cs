// vybe-test: csharp/csharp_deconstruction/deconstruct_method_returns_three_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var color = new Color(1, 2, 3);
var (red, green, blue) = color;
__P((red + green + blue).ToString());
__Check("6");

class Color {
    int r;
    int g;
    int b;
    public Color(int r, int g, int b) { this.r = r; this.g = g; this.b = b; }
    public void Deconstruct(out int red, out int green, out int blue) {
        red = r;
        green = g;
        blue = b;
    }
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
