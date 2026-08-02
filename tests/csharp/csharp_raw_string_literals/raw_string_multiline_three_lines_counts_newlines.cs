// vybe-test: csharp/csharp_raw_string_literals/raw_string_multiline_three_lines_counts_newlines
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""one
two
three"""; __Check((text.Split('\n').Length).ToString(), "3");
