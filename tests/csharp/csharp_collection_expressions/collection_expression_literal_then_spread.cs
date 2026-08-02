// vybe-test: csharp/csharp_collection_expressions/collection_expression_literal_then_spread
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] rest = [2, 3];
int[] all = [1, ..rest];
__Check((all.Length).ToString(), "3");
