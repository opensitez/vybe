// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_struct_get_only
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Row { int[] cells = { 1, 2, 3 }; public int this[int c] => cells[c]; }
__Check((new Row()[0]).ToString(), "1");
