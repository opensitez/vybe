// vybe-test: csharp/csharp_with_expression_records/with_string_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Tag(string Name); var n=(new Tag("old")) with{Name="new"}; __Check((n.Name).ToString(), "new");
