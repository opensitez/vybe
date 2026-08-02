// vybe-test: csharp/oop_advanced/operator_overload_plus
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Vector {
    public double X { get; set; }
    public double Y { get; set; }
    public Vector(double x, double y) { X = x; Y = y; }
    public static Vector operator +(Vector a, Vector b) {
        return new Vector(a.X + b.X, a.Y + b.Y);
    }
}
var a = new Vector(1, 2);
var b = new Vector(3, 4);
var c = a + b;
__Check((c.X).ToString(), "4");
__Check((c.Y).ToString(), "6");
