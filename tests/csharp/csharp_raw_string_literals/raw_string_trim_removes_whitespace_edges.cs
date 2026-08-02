// vybe-test: csharp/csharp_raw_string_literals/raw_string_trim_removes_whitespace_edges
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""  trim  """; __Check((text.Trim()).ToString(), "trim");
