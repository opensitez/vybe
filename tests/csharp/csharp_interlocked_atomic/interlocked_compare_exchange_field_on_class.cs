// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public int Value = 0;
    public int Cas(int expected, int desired) {
        return System.Threading.Interlocked.CompareExchange(ref Value, desired, expected);
    }
}
var c = new Counter();
__Check((c.Cas(0, 11)).ToString(), "0");
__Check((c.Value).ToString(), "11");
