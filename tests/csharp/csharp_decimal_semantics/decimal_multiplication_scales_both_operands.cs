// vybe-test: csharp/csharp_decimal_semantics/decimal_multiplication_scales_both_operands
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal rate = 1.5m; decimal hours = 2m; __Check((rate * hours).ToString(), "3.0");
