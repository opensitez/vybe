// vybe-test: csharp/csharp_properties/expression_bodied_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rect { public int W,H; public int Area => W * H; }
__Check((new Rect{W=3,H=4}.Area).ToString(), "12");
