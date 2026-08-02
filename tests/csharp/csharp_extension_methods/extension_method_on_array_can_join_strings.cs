// vybe-test: csharp/csharp_extension_methods/extension_method_on_array_can_join_strings
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class StringArrayExt { public static string JoinAll(this string[] values) { return string.Join(",", values); } } } __Check((new[] { "a", "b" }.JoinAll()).ToString(), "a,b");
