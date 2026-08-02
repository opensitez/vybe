// vybe-test: csharp/csharp_with_expression_records/with_positional_preserves_string
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(string S,int N); var q=(new Pair("x",1)) with{N=2}; __Check((q.S).ToString(), "x");
