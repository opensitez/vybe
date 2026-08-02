// vybe-test: csharp/csharp_raw_string_literals/raw_string_substring_extracts_suffix
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""abcdef"""; __Check((text.Substring(4)).ToString(), "ef");
