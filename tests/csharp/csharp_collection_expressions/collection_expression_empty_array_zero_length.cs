// vybe-test: csharp/csharp_collection_expressions/collection_expression_empty_array_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] empty = [];
__Check((empty.Length).ToString(), "0");
