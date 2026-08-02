// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_wrong_expected
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 12;
var prev = System.Threading.Interlocked.CompareExchange(ref slot, 99, 0);
__Check((prev).ToString(), "12");
__Check((slot).ToString(), "12");
