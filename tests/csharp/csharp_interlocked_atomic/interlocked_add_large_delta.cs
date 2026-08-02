// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_large_delta
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 100;
__Check((System.Threading.Interlocked.Add(ref total, 900)).ToString(), "1000");
__Check((total).ToString(), "1000");
