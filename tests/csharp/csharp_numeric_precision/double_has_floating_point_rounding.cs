// vybe-test: csharp/csharp_numeric_precision/double_has_floating_point_rounding
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double a=0.1, b=0.2;
__Check((a+b==0.3).ToString(), "False");
