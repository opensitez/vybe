// vybe-test: csharp/csharp_strings_ext/string_padleft_padright
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("5".PadLeft(3, '0')).ToString(), "005");
__Check(("5".PadRight(3, '0')).ToString(), "500");
