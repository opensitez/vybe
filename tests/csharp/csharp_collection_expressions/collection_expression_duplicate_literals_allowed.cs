// vybe-test: csharp/csharp_collection_expressions/collection_expression_duplicate_literals_allowed
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = [7, 7, 7];
__Check((arr[2]).ToString(), "7");
