// vybe-test: csharp/csharp_string_interpolation/alignment_specifier_pads_right_aligned
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(($"{"x",5}").ToString(), "    x");
