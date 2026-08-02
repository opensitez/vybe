// vybe-test: csharp/csharp_with_expression_records/with_tostring_after
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Tag(string Name); var u=(new Tag("a")) with{Name="b"}; __Check((u.ToString().Contains("b")).ToString(), "True");
