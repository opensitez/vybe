// vybe-test: csharp/csharp_collection_expressions/collection_expression_with_zero_values
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] zeros = [0, 0, 0];
__Check((zeros[1]).ToString(), "0");
