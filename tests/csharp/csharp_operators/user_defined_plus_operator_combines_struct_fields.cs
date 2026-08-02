// vybe-test: csharp/csharp_operators/user_defined_plus_operator_combines_struct_fields
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec2 {
    public int X;
    public int Y;
    public static Vec2 operator +(Vec2 a, Vec2 b) =>
        new Vec2 { X = a.X + b.X, Y = a.Y + b.Y };
}
var sum = new Vec2 { X = 1, Y = 2 } + new Vec2 { X = 3, Y = 4 };
__Check((sum.X).ToString(), "4");
__Check((sum.Y).ToString(), "6");
