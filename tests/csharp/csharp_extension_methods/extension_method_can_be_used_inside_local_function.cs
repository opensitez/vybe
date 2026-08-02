// vybe-test: csharp/csharp_extension_methods/extension_method_can_be_used_inside_local_function
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class NumberExt { public static int Inc(this int value) { return value + 1; } } } int Read() { return 9.Inc(); } __Check((Read()).ToString(), "10");
