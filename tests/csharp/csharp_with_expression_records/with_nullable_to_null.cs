// vybe-test: csharp/csharp_with_expression_records/with_nullable_to_null
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Maybe(int? N); var z=(new Maybe(5)) with{N=null}; __Check((z.N.HasValue).ToString(), "False");
