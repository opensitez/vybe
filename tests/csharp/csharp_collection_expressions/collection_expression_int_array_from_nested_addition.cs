// vybe-test: csharp/csharp_collection_expressions/collection_expression_int_array_from_nested_addition
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = [1 + 2, 3 + 4];
__Check((arr[1]).ToString(), "7");
