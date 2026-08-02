// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_two_int_params
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Grid { int[,] m = { { 1, 2 }, { 3, 4 } }; public int this[int r, int c] => m[r, c]; }
__Check((new Grid()[1, 0]).ToString(), "3");
