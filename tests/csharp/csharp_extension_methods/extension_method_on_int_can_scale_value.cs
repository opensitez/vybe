// vybe-test: csharp/csharp_extension_methods/extension_method_on_int_can_scale_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class NumberExt { public static int Triple(this int value) { return value * 3; } } } __Check((4.Triple()).ToString(), "12");
