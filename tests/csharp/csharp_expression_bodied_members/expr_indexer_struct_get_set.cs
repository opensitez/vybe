// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_struct_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct PairStore { int a, b; public int this[int slot] { get => slot == 0 ? a : b; set { if (slot == 0) a = value; else b = value; } } }
var p = new PairStore(); p[0] = 3; p[1] = 9; __Check((p[0]).ToString(), "3"); __Check((p[1]).ToString(), "9");
