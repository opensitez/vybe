// vybe-test: csharp/csharp_raw_string_literals/raw_string_insert_adds_middle_text
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""ac"""; __Check((text.Insert(1,"b")).ToString(), "abc");
