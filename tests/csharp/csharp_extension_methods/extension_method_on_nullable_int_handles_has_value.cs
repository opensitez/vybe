// vybe-test: csharp/csharp_extension_methods/extension_method_on_nullable_int_handles_has_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class NullableExt { public static string Describe(this int? value) { return value.HasValue ? value.Value.ToString() : "none"; } } } int? value = 8; __Check((value.Describe()).ToString(), "8");
