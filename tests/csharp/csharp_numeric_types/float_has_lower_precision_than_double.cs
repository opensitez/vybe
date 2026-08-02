// vybe-test: csharp/csharp_numeric_types/float_has_lower_precision_than_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((sizeof(float)).ToString(), "4"); __Check((sizeof(double)).ToString(), "8");
