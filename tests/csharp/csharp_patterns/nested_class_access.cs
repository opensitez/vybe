// vybe-test: csharp/csharp_patterns/nested_class_access
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    public int Value = 10;
    public class Inner {
        public int Value = 20;
    }
}
var o = new Outer();
var i = new Outer.Inner();
__Check((o.Value).ToString(), "10");
__Check((i.Value).ToString(), "20");
