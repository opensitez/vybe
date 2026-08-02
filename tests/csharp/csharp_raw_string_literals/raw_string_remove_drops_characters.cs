// vybe-test: csharp/csharp_raw_string_literals/raw_string_remove_drops_characters
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""abcde"""; __Check((text.Remove(1,2)).ToString(), "ade");
