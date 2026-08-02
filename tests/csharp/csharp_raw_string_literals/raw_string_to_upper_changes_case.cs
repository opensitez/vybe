// vybe-test: csharp/csharp_raw_string_literals/raw_string_to_upper_changes_case
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""abc"""; __Check((text.ToUpper()).ToString(), "ABC");
