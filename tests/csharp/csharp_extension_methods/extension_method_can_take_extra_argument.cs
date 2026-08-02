// vybe-test: csharp/csharp_extension_methods/extension_method_can_take_extra_argument
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class TextExt { public static string Wrap(this string value, string prefix) { return prefix + value; } } } __Check(("core".Wrap("pre-")).ToString(), "pre-core");
