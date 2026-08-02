// vybe-test: csharp/csharp_extension_methods/extension_method_can_return_tuple_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class TextExt { public static (string, int) Pack(this string value) { return (value, value.Length); } } } var result = "tool".Pack(); __Check((result.Item1).ToString(), "tool"); __Check((result.Item2).ToString(), "4");
