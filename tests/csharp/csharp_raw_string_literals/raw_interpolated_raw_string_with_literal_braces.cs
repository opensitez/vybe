// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_raw_string_with_literal_braces
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=1; string text=$"""value={n} end"""; __Check((text.EndsWith(" end")).ToString(), "True");
