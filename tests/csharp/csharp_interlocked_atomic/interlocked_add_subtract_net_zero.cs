// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_subtract_net_zero
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 50;
System.Threading.Interlocked.Add(ref total, 10);
System.Threading.Interlocked.Add(ref total, -10);
__Check((total).ToString(), "50");
