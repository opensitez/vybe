// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_multiple_steps_sum
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 0;
System.Threading.Interlocked.Add(ref total, 2);
System.Threading.Interlocked.Add(ref total, 3);
System.Threading.Interlocked.Add(ref total, 5);
__Check((total).ToString(), "10");
