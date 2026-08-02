// vybe-test: csharp/csharp_equality_contracts/struct_equality_compares_field_values_when_equals_overridden
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point {
    public int X;
    public int Y;
    public bool Equals(Point other) { return X == other.X && Y == other.Y; }
}
var left = new Point { X = 2, Y = 3 };
var right = new Point { X = 2, Y = 3 };
__Check((left.Equals(right)).ToString(), "True");
