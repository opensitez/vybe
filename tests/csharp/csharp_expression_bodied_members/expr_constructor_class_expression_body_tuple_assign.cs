// vybe-test: csharp/csharp_expression_bodied_members/expr_constructor_class_expression_body_tuple_assign
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X, Y; public Point(int x, int y) => (X, Y) = (x, y); }
var p = new Point(3, 4); __Check((p.X).ToString(), "3"); __Check((p.Y).ToString(), "4");
