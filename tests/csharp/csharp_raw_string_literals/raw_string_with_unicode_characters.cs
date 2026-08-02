// vybe-test: csharp/csharp_raw_string_literals/raw_string_with_unicode_characters
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""café"""; __Check((text.Contains("é")).ToString(), "True");
