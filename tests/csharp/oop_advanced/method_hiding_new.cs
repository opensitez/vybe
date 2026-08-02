// vybe-test: csharp/oop_advanced/method_hiding_new
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public string Speak() { return "base"; }
}
class Child : Base {
    public new string Speak() { return "child"; }
}
var c = new Child();
__Check((c.Speak()).ToString(), "child");
