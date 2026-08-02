// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_multiline_preserves_newline_and_value
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int id=9; string text=$"""id:
{id}"""; __Check((text.Contains("9")).ToString(), "True");
