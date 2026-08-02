// vybe-test: csharp/csharp_collection_expressions/collection_expression_array_first_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = [7, 8, 9];
__Check((arr[0]).ToString(), "7");
