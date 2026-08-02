// vybe-test: csharp/csharp_with_expression_records/with_zero_int
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int N); var z=(new V(5)) with{N=0}; __Check((z.N).ToString(), "0");
