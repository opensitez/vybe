// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_empty_array_adds_nothing
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] empty = [];
int[] all = [1, ..empty, 2];
__Check((all.Length).ToString(), "2");
