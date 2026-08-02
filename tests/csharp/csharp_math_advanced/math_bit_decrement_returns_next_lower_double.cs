// vybe-test: csharp/csharp_math_advanced/math_bit_decrement_returns_next_lower_double
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.BitDecrement(1.0)<1.0).ToString(), "True");
