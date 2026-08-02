// vybe-test: csharp/csharp_patterns/factory_method
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shape {
    public string Type;
    private Shape(string t) { Type = t; }
    public static Shape Circle() { return new Shape("circle"); }
    public static Shape Square() { return new Shape("square"); }
}
var c = Shape.Circle();
var s = Shape.Square();
__Check((c.Type).ToString(), "circle");
__Check((s.Type).ToString(), "square");
