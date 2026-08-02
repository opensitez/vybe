// vybe-test: csharp/csharp_with_expression_records/with_mutable_separate_instance
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Box{public int V{get;set;}} var b=(new Box{V=1}) with{V=2}; __Check((b.V).ToString(), "2");
