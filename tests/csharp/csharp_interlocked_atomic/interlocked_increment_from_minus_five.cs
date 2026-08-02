// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_from_minus_five
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = -5;
__Check((System.Threading.Interlocked.Increment(ref counter)).ToString(), "-4");
__Check((counter).ToString(), "-4");
