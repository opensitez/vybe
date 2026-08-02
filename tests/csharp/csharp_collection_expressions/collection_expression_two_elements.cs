// vybe-test: csharp/csharp_collection_expressions/collection_expression_two_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] pair = [4, 5];
__Check((pair[0] + pair[1]).ToString(), "9");
