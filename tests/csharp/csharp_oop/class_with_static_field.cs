// vybe-test: csharp/csharp_oop/class_with_static_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public static int Count = 0;
    public Counter() { Count++; }
}
var a = new Counter();
var b = new Counter();
var c = new Counter();
__Check((Counter.Count).ToString(), "3");
