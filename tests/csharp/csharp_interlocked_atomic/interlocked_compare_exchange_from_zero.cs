// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_from_zero
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 0;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 5, 0);
__Check((previous).ToString(), "0");
__Check((slot).ToString(), "5");
