// vybe-test: csharp/csharp_with_expression_records/with_max_int
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int N); var m=(new V(1)) with{N=int.MaxValue}; __Check((m.N==int.MaxValue).ToString(), "True");
