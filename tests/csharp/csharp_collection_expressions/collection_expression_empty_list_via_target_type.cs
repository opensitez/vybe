// vybe-test: csharp/csharp_collection_expressions/collection_expression_empty_list_via_target_type
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> list = [];
__Check((list.Count).ToString(), "0");
