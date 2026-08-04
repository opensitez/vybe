// vybe-test: csharp/csharp_deconstruction/deconstruct_method_returns_three_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

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
__P((red + green + blue).ToString());
__Check("6");
