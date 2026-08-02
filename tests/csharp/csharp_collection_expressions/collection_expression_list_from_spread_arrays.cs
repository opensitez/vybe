// vybe-test: csharp/csharp_collection_expressions/collection_expression_list_from_spread_arrays
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = [1, 2];
int[] b = [3];
System.Collections.Generic.List<int> list = [..a, ..b];
__Check((list.Count).ToString(), "3"); __Check((list[2]).ToString(), "3");
