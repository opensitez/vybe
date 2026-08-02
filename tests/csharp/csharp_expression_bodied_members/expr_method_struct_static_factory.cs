// vybe-test: csharp/csharp_expression_bodied_members/expr_method_struct_static_factory
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X, Y; public static Point Origin() => new Point { X = 0, Y = 0 }; }
var p = Point.Origin();
__Check((p.X).ToString(), "0"); __Check((p.Y).ToString(), "0");
