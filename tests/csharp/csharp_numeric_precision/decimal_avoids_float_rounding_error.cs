// vybe-test: csharp/csharp_numeric_precision/decimal_avoids_float_rounding_error
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal a=0.1m, b=0.2m;
__Check((a+b==0.3m).ToString(), "True");
