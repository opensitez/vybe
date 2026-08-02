// vybe-test: csharp/csharp_deconstruction_patterns/custom_class_with_deconstruct_method_supports_deconstruction
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Size {
    public int W, H;
    public void Deconstruct(out int w, out int h) { w=W; h=H; }
}
var (w, h) = new Size{W=3,H=4};
__Check((w).ToString(), "3"); __Check((h).ToString(), "4");
