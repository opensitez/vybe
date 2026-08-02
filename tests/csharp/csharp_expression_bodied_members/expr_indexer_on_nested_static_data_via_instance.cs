// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_on_nested_static_data_via_instance
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Lookup { int[] table = { 5, 6, 7 }; public int this[int i] => table[i]; }
__Check((new Lookup()[2]).ToString(), "7");
