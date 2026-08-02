// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_overwrites_existing
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 77;
__Check((System.Threading.Interlocked.Exchange(ref slot, 3)).ToString(), "77");
__Check((slot).ToString(), "3");
