// vybe-test: csharp/csharp_interlocked_atomic/interlocked_decrement_to_zero
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 1;
__Check((System.Threading.Interlocked.Decrement(ref counter)).ToString(), "0");
__Check((counter).ToString(), "0");
