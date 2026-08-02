// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_in_middle
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] mid = [2, 3];
int[] all = [1, ..mid, 4];
__Check((all[1]).ToString(), "2"); __Check((all[3]).ToString(), "4");
