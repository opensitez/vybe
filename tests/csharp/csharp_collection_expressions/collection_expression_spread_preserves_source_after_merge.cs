// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_preserves_source_after_merge
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = [1, 2];
int[] b = [..a, 3];
__Check((a.Length).ToString(), "2"); __Check((b.Length).ToString(), "3");
