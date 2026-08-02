// vybe-test: csharp/csharp_extension_methods/generic_extension_method_returns_same_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class EchoExt { public static T Echo<T>(this T value) { return value; } } } __Check(("hi".Echo()).ToString(), "hi");
