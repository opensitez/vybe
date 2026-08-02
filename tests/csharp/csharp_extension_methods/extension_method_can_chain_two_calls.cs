// vybe-test: csharp/csharp_extension_methods/extension_method_can_chain_two_calls
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class TextExt { public static string AddA(this string value) { return value + "a"; } public static string AddB(this string value) { return value + "b"; } } } __Check(("x".AddA().AddB()).ToString(), "xab");
