// vybe-test: csharp/csharp_decimal_semantics/decimal_mixed_addition_with_int_promotes_int_operand
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal baseAmount = 2.5m; __Check((baseAmount + 2).ToString(), "4.5");
