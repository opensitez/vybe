// vybe-test: csharp/csharp_indexers/readonly_indexer_exposes_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Odds{public int this[int n]=>2*n+1;}
__Check((new Odds()[4]).ToString(), "9");
