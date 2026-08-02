// vybe-test: csharp/csharp_with_expression_records/with_double_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Rate(double V); var s=(new Rate(1.1)) with{V=2.2}; __Check((s.V).ToString(), "2.2");
