// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_then_increment
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 2;
System.Threading.Interlocked.Exchange(ref slot, 10);
__Check((System.Threading.Interlocked.Increment(ref slot)).ToString(), "11");
__Check((slot).ToString(), "11");
