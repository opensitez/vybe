// vybe-test: csharp/csharp_raw_string_literals/raw_string_line_count_in_multiline_literal
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""row1
row2
row3
row4"""; __Check((text.Split('\n').Length).ToString(), "4");
