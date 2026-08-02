// vybe-test: csharp/csharp_extension_methods/extension_method_on_object_accepts_boxed_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class ObjectExt { public static string Kind(this object value) { return value.GetType().Name; } } } object value = 3; __Check((value.Kind()).ToString(), "Int32");
