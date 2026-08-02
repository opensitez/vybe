// vybe-test: csharp/csharp_with_expression_records/with_negative_int
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int N); var n=(new V(5)) with{N=-1}; __Check((n.N).ToString(), "-1");
