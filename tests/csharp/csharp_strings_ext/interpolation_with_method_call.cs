// vybe-test: csharp/csharp_strings_ext/interpolation_with_method_call
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name = "world";
__Check(($"Hello {name.ToUpper()}!").ToString(), "Hello WORLD!");
