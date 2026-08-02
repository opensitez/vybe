// vybe-test: csharp/csharp_math_functions/math_min_returns_smaller_of_two_values
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Min(3, 7)).ToString(), "3");
