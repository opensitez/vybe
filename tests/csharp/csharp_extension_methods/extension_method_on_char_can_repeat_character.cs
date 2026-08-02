// vybe-test: csharp/csharp_extension_methods/extension_method_on_char_can_repeat_character
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public static class CharExt { public static string Repeat(this char value, int count) { return new string(value, count); } } } __Check(('x'.Repeat(3)).ToString(), "xxx");
