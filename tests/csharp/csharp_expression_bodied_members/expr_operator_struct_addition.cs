// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_struct_addition
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec2 { public int X, Y; public static Vec2 operator +(Vec2 a, Vec2 b) => new Vec2 { X = a.X + b.X, Y = a.Y + b.Y }; }
var v = new Vec2 { X = 1, Y = 2 } + new Vec2 { X = 3, Y = 4 };
__Check((v.X).ToString(), "4"); __Check((v.Y).ToString(), "6");
