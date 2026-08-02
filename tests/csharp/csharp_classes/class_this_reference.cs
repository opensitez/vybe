// vybe-test: csharp/csharp_classes/class_this_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    private int count = 0;
    public void Increment() { this.count++; }
    public int GetCount() { return this.count; }
}
var c = new Counter();
c.Increment();
c.Increment();
c.Increment();
__Check((c.GetCount()).ToString(), "3");
