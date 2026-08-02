// vybe-test: csharp/csharp_extension_methods/extension_method_can_use_namespace_import_from_nested_namespace
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo.Tools; namespace Demo.Tools { public static class TextExt { public static string Bang(this string value) { return value + "!"; } } } __Check(("go".Bang()).ToString(), "go!");
