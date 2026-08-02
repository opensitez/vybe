// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_into_new_list_count
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = [5, 6, 7];
System.Collections.Generic.List<int> list = [..data];
__Check((list[2]).ToString(), "7");
