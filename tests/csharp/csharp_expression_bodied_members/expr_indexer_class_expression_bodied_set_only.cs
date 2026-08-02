// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_expression_bodied_set_only
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Store { int v; public int Value { get { return v; } set => v = value; } }
var s = new Store(); s.Value = 11; __Check((s.Value).ToString(), "11");
