// vybe-test: csharp/csharp_collection_expressions/collection_expression_returned_from_method
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] Make() => [4, 5, 6];
__Check((Make()[1]).ToString(), "5");
