// vybe-test: csharp/csharp_nameof_expressions/nameof_indexer_declaring_type_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag{public int this[int i]{get=>i;set{}}} __Check((nameof(Bag)).ToString(), "Bag");
