// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_returns_each_previous
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 4;
int old1 = System.Threading.Interlocked.Exchange(ref slot, 5);
int old2 = System.Threading.Interlocked.Exchange(ref slot, 6);
__Check((old1 + old2).ToString(), "9");
__Check((slot).ToString(), "6");
