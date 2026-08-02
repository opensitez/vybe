// vybe-test: csharp/csharp_collection_expressions/collection_expression_target_typed_list_add_after
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> list = [1, 2];
list.Add(3);
__Check((list[2]).ToString(), "3");
