// vybe-test: csharp/csharp_math_advanced/math_log2_of_power_of_two
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((int)System.Math.Log2(8)).ToString(), "3");
