// vybe-test: csharp/csharp_raw_string_literals/raw_string_multiline_preserves_internal_newline
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""line1
line2"""; __Check((text.Contains("\n")).ToString(), "True");
