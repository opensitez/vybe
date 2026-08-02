// vybe-test: csharp/csharp_decimal_semantics/decimal_cast_from_int_promotes_to_integral_decimal
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal value = (decimal)7; __Check((value).ToString(), "7");
