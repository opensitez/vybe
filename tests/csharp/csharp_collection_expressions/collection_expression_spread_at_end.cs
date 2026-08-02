// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_at_end
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] head = [1, 2];
int[] all = [9, ..head];
__Check((all[0]).ToString(), "9"); __Check((all[2]).ToString(), "2");
