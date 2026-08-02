// vybe-test: csharp/csharp_decimal_semantics/decimal_subtraction_yields_exact_difference_for_currency_style_values
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal price = 19.99m; decimal discount = 4.50m; __Check((price - discount).ToString(), "15.49");
