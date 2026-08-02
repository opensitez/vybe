// vybe-test: csharp/csharp_patterns/method_overloading
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Printer {
    public string Print(int x) { return "int:" + x; }
    public string Print(string x) { return "str:" + x; }
    public string Print(int x, int y) { return "pair:" + x + "," + y; }
}
var p = new Printer();
__Check((p.Print(42)).ToString(), "int:42");
__Check((p.Print("hi")).ToString(), "str:hi");
__Check((p.Print(1, 2)).ToString(), "pair:1,2");
