// vybe-test: csharp/csharp_decimal_financial/decimal_financial_multiply_unit_price_by_qty
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal unit=12.75m; int qty=3; __Check((unit*qty).ToString(), "38.25");
