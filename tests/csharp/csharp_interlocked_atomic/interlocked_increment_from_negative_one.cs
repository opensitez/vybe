// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_from_negative_one
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = -1;
__Check((System.Threading.Interlocked.Increment(ref counter)).ToString(), "0");
__Check((counter).ToString(), "0");
