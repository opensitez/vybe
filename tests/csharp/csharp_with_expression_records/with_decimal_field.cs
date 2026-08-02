// vybe-test: csharp/csharp_with_expression_records/with_decimal_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Price(decimal A); var q=(new Price(1.5m)) with{A=9.99m}; __Check((q.A).ToString(), "9.99");
