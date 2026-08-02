// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public int Value = 0;
    public void Bump() { System.Threading.Interlocked.Increment(ref Value); }
}
var c = new Counter();
c.Bump();
c.Bump();
__Check((c.Value).ToString(), "2");
