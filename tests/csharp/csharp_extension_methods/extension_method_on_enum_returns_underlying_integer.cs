// vybe-test: csharp/csharp_extension_methods/extension_method_on_enum_returns_underlying_integer
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; enum Mode { Off = 0, On = 2 } namespace Demo { public static class ModeExt { public static int Code(this Mode mode) { return (int)mode; } } } __Check((Mode.On.Code()).ToString(), "2");
