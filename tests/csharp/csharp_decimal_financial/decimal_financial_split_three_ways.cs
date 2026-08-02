// vybe-test: csharp/csharp_decimal_financial/decimal_financial_split_three_ways
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal total=10.00m; __Check((total/3m>3.33m&&total/3m<3.34m).ToString(), "True");
