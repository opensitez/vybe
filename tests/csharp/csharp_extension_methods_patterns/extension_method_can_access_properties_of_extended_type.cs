// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_can_access_properties_of_extended_type
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Width, Height; }
static class BoxExt { public static int Area(this Box b) => b.Width*b.Height; }
__Check((new Box{Width=3,Height=4}.Area()).ToString(), "12");
