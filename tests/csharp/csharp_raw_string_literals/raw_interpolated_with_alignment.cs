// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_with_alignment
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=42; string text=$"""{n,5}"""; __Check((text.Trim().Length>=2).ToString(), "True");
