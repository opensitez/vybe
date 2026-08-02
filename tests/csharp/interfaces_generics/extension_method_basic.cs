// vybe-test: csharp/interfaces_generics/extension_method_basic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class StringExtensions {
    public static string Reverse(this string s) {
        char[] chars = s.ToCharArray();
        Array.Reverse(chars);
        return new string(chars);
    }
}
__Check(("hello".Reverse()).ToString(), "olleh");
