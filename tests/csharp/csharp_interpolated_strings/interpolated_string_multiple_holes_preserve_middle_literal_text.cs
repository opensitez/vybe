// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_multiple_holes_preserve_middle_literal_text
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(($"a{1}b{2}c").ToString(), "a1b2c");
