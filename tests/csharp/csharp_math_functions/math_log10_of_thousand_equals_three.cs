// vybe-test: csharp/csharp_math_functions/math_log10_of_thousand_equals_three
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Log10(1000)).ToString(), "3");
