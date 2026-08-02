// vybe-test: csharp/csharp_decimal_financial/decimal_financial_divide_per_unit_cost
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal bill=100.00m; decimal seats=6m; __Check((bill/seats>16.6m&&bill/seats<16.7m).ToString(), "True");
