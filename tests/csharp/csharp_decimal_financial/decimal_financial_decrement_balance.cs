// vybe-test: csharp/csharp_decimal_financial/decimal_financial_decrement_balance
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal balance=5.00m; balance-=0.01m; __Check((balance).ToString(), "4.99");
