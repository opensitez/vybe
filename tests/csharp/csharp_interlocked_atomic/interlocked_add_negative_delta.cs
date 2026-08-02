// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_negative_delta
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 20;
__Check((System.Threading.Interlocked.Add(ref total, -5)).ToString(), "15");
__Check((total).ToString(), "15");
