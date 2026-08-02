// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_string_chains_after_built_in_method
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class StrExt { public static string Shout(this string s) => s.ToUpper() + "!"; }
__Check(("hello".Shout()).ToString(), "HELLO!");
