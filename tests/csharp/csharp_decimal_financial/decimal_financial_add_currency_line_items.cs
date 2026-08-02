// vybe-test: csharp/csharp_decimal_financial/decimal_financial_add_currency_line_items
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal subtotal=19.99m+4.50m+0.01m; __Check((subtotal).ToString(), "24.50");
