// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_get_set_both_expression
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { int _v; public int Value { get => _v; set => _v = value; } }
var b = new Box(); b.Value = 9; __Check((b.Value).ToString(), "9");
