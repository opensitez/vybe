// vybe-test: csharp/csharp_raw_string_literals/raw_string_embedded_double_quotes_without_backslash
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""say "hi" now"""; __Check((text.Contains(""hi"")).ToString(), "True");
