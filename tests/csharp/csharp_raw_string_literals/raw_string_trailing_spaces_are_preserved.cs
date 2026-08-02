// vybe-test: csharp/csharp_raw_string_literals/raw_string_trailing_spaces_are_preserved
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""trail  """; __Check((text.EndsWith("  ")).ToString(), "True");
