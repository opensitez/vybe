// vybe-test: csharp/csharp_collection_expressions/collection_expression_readonly_span_length
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<int> s = [9, 8, 7];
__Check((s.Length).ToString(), "3");
