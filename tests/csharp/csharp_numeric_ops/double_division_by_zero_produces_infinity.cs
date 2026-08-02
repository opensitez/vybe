// vybe-test: csharp/csharp_numeric_ops/double_division_by_zero_produces_infinity
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d=1.0/0.0;
__Check((double.IsInfinity(d)).ToString(), "True");
