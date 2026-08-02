// vybe-test: csharp/csharp_expression_bodied_members/expr_property_struct_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Slot { int _n; public int N { get => _n; set => _n = value; } }
var s = new Slot(); s.N = 7; __Check((s.N).ToString(), "7");
