// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_negative_values
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = -2;
var prev = System.Threading.Interlocked.CompareExchange(ref slot, -1, -2);
__Check((prev).ToString(), "-2");
__Check((slot).ToString(), "-1");
