// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_with_zero_new
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 55;
__Check((System.Threading.Interlocked.Exchange(ref slot, 0)).ToString(), "55");
__Check((slot).ToString(), "0");
