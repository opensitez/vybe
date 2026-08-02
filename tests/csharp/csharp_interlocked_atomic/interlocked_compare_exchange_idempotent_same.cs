// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_idempotent_same
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 3;
var p1 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
var p2 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
__Check((p1 + p2).ToString(), "6");
__Check((slot).ToString(), "8");
