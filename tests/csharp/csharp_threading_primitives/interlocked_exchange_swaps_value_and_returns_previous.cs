// vybe-test: csharp/csharp_threading_primitives/interlocked_exchange_swaps_value_and_returns_previous
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 1;
__Check((System.Threading.Interlocked.Exchange(ref slot, 9)).ToString(), "1");
__Check((slot).ToString(), "9");
