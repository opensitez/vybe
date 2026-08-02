// vybe-test: csharp/csharp_threading_primitives/interlocked_compare_exchange_updates_only_when_current_matches
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 7;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 99, 7);
__Check((previous).ToString(), "7");
__Check((slot).ToString(), "99");
