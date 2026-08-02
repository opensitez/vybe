// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_returns_old_and_sets_new
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 1;
__Check((System.Threading.Interlocked.Exchange(ref slot, 9)).ToString(), "1");
__Check((slot).ToString(), "9");
