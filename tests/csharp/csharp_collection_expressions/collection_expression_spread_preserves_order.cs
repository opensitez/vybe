// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_preserves_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = [1, 2];
int[] b = [3];
int[] c = [..a, ..b];
__Check((c[0]).ToString(), "1"); __Check((c[2]).ToString(), "3");
