// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_single_element_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] one = [42];
int[] two = [..one, 99];
__Check((two.Length).ToString(), "2"); __Check((two[1]).ToString(), "99");
