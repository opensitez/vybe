// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_copy_via_self
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src = [1, 2];
int[] copy = [..src];
__Check((copy[1]).ToString(), "2");
