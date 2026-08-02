// vybe-test: csharp/csharp_expression_bodied/expression_bodied_indexer
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag{int[]data={1,2,3};public int this[int i]=>data[i];}
__Check((new Bag()[2]).ToString(), "3");
