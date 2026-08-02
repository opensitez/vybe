// vybe-test: csharp/csharp_collection_expressions/collection_expression_list_int_middle_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> list = [10, 20, 30];
__Check((list[1]).ToString(), "20");
