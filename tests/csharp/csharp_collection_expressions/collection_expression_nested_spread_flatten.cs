// vybe-test: csharp/csharp_collection_expressions/collection_expression_nested_spread_flatten
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = [1]; int[] b = [2]; int[] c = [3];
int[] all = [..a, ..b, ..c];
__Check((string.Join(",", all)).ToString(), "1,2,3");
