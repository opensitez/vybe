// vybe-test: csharp/csharp_with_expression_records/with_long_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Wide(long V); var x=(new Wide(10L)) with{V=20L}; __Check((x.V).ToString(), "20");
