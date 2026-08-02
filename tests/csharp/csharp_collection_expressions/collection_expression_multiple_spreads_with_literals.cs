// vybe-test: csharp/csharp_collection_expressions/collection_expression_multiple_spreads_with_literals
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = [1, 2]; int[] b = [3];
int[] c = [0, ..a, ..b, 4];
__Check((c[0]).ToString(), "0"); __Check((c[4]).ToString(), "4");
