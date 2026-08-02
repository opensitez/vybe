// vybe-test: csharp/csharp_extension_methods/extension_method_on_string_returns_length_label
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class TextExt { public static string Label(this string value) { return value + ":" + value.Length; } } } __Check(("abc".Label()).ToString(), "abc:3");
