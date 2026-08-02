// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public int Value = 5;
    public void Add(int n) { System.Threading.Interlocked.Add(ref Value, n); }
}
var c = new Counter();
c.Add(3);
__Check((c.Value).ToString(), "8");
