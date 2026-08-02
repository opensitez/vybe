// vybe-test: csharp/oop_advanced/expression_bodied_method
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc {
    public int Square(int x) => x * x;
    public string Greet(string name) => "Hello " + name;
}
var c = new Calc();
__Check((c.Square(7)).ToString(), "49");
__Check((c.Greet("World")).ToString(), "Hello World");
