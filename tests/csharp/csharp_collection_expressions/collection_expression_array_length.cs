// vybe-test: csharp/csharp_collection_expressions/collection_expression_array_length
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = [10, 20, 30];
__Check((arr.Length).ToString(), "3");
