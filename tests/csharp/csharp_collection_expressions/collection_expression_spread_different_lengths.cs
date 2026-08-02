// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_different_lengths
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] small = [1];
int[] big = [2, 3, 4, 5];
int[] all = [..small, ..big];
__Check((all.Length).ToString(), "5");
