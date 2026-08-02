// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_from_zero_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 0;
__Check((System.Threading.Interlocked.Increment(ref counter)).ToString(), "1");
__Check((counter).ToString(), "1");
