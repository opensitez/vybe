// vybe-test: csharp/csharp_struct_features/struct_with_custom_constructor_sets_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Rect { public int W,H; public Rect(int w, int h) { W=w; H=h; } }
var r = new Rect(3, 4);
__Check((r.W * r.H).ToString(), "12");
