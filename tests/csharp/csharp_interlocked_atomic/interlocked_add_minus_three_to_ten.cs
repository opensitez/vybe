// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_minus_three_to_ten
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 10;
__Check((System.Threading.Interlocked.Add(ref total, -3)).ToString(), "7");
__Check((total).ToString(), "7");
