// vybe-test: csharp/csharp_math_functions/math_truncate_removes_fractional_part_toward_zero
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Truncate(9.9)).ToString(), "9");
