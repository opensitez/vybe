// vybe-test: csharp/csharp_raw_string_literals/raw_string_backslash_is_literal_not_escape
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""C:\temp\file"""; __Check((text.Contains(@"")).ToString(), "True");
