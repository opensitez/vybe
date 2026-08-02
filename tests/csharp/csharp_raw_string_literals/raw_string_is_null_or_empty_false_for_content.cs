// vybe-test: csharp/csharp_raw_string_literals/raw_string_is_null_or_empty_false_for_content
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""x"""; __Check((string.IsNullOrEmpty(text)).ToString(), "False");
