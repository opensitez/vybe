// vybe-test: csharp/csharp_deconstruction/deconstruct_method_returns_three_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

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
var color = new Color(1, 2, 3);
var (red, green, blue) = color;
__Check((red + green + blue).ToString(), "6");
