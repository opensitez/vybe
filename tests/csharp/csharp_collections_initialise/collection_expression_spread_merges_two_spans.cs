// vybe-test: csharp/csharp_collections_initialise/collection_expression_spread_merges_two_spans
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a=[1,2,3];
int[] b=[4,5,6];
int[] c=[..a,..b];
__Check((c.Length).ToString(), "6"); __Check((c[3]).ToString(), "4");
