// vybe-test: csharp/csharp_collection_expressions/collection_expression_mixed_spread_and_literal_sum
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] mid = [2, 3];
int[] all = [1, ..mid, 4];
__Check((all[0] + all[3]).ToString(), "5");
