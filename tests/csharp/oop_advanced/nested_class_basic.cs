// vybe-test: csharp/oop_advanced/nested_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    public class Inner {
        public string Hello() { return "inner"; }
    }
}
var i = new Outer.Inner();
__Check((i.Hello()).ToString(), "inner");
