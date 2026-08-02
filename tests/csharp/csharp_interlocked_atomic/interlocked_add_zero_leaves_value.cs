// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_zero_leaves_value
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 8;
__Check((System.Threading.Interlocked.Add(ref total, 0)).ToString(), "8");
__Check((total).ToString(), "8");
