// vybe-test: csharp/csharp_raw_string_literals/raw_string_literal_newline_not_escape_sequence
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""top
bottom"""; __Check((text.IndexOf('\n')>0).ToString(), "True");
