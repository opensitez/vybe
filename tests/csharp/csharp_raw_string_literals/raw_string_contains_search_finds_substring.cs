// vybe-test: csharp/csharp_raw_string_literals/raw_string_contains_search_finds_substring
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""hello world"""; __Check((text.Contains("world")).ToString(), "True");
