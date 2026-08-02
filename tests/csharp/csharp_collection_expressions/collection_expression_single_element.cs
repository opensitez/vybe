// vybe-test: csharp/csharp_collection_expressions/collection_expression_single_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] one = [99];
__Check((one[0]).ToString(), "99");
