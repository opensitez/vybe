// vybe-test: csharp/csharp_numeric_ops/long_arithmetic_handles_large_values_exactly
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long a=3_000_000_000L; long b=a*2;
__Check((b).ToString(), "6000000000");
