// vybe-test: csharp/csharp_collections_initialise/span_collection_expression_works_with_stack_alloc_semantics
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> s=[1,2,3];
__Check((s.Length).ToString(), "3");
