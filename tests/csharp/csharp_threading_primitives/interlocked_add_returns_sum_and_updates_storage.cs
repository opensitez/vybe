// vybe-test: csharp/csharp_threading_primitives/interlocked_add_returns_sum_and_updates_storage
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 10;
__Check((System.Threading.Interlocked.Add(ref total, 4)).ToString(), "14");
__Check((total).ToString(), "14");
