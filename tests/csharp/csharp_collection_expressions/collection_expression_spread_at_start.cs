// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_at_start
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] tail = [3, 4];
int[] all = [..tail, 1, 2];
__Check((all[0]).ToString(), "3"); __Check((all[2]).ToString(), "1");
