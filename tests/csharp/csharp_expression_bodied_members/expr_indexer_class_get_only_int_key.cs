// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_get_only_int_key
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag { int[] data = { 10, 20, 30 }; public int this[int i] => data[i]; }
__Check((new Bag()[1]).ToString(), "20");
