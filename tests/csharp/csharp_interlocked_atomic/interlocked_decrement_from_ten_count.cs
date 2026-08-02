// vybe-test: csharp/csharp_interlocked_atomic/interlocked_decrement_from_ten_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 10;
__Check((System.Threading.Interlocked.Decrement(ref counter)).ToString(), "9");
__Check((counter).ToString(), "9");
