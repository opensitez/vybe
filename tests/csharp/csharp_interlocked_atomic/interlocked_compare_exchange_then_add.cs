// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_then_add
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 4;
System.Threading.Interlocked.CompareExchange(ref slot, 4, 4);
__Check((System.Threading.Interlocked.Add(ref slot, 6)).ToString(), "10");
__Check((slot).ToString(), "10");
