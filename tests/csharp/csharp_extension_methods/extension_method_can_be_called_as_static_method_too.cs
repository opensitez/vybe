// vybe-test: csharp/csharp_extension_methods/extension_method_can_be_called_as_static_method_too
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class TextExt { public static string Wrap(this string value) { return "[" + value + "]"; } } } __Check((TextExt.Wrap("x")).ToString(), "[x]");
