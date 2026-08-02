// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_then_decrement_net_zero
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 0;
System.Threading.Interlocked.Increment(ref counter);
System.Threading.Interlocked.Decrement(ref counter);
__Check((counter).ToString(), "0");
