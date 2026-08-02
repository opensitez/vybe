// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_list_into_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> list = new() { 1, 2 };
int[] arr = [..list, 3];
__Check((arr[2]).ToString(), "3");
