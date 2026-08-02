// vybe-test: csharp/csharp_collection_expressions/collection_expression_index_access_second
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = [10, 11, 12];
__Check((arr[1]).ToString(), "11");
