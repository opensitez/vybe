// vybe-test: csharp/csharp_collection_expressions/collection_expression_span_first_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> s = [5, 6];
__Check((s[0]).ToString(), "5");
