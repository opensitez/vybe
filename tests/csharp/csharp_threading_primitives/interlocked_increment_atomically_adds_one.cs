// vybe-test: csharp/csharp_threading_primitives/interlocked_increment_atomically_adds_one
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 5;
__Check((System.Threading.Interlocked.Increment(ref count)).ToString(), "6");
__Check((count).ToString(), "6");
