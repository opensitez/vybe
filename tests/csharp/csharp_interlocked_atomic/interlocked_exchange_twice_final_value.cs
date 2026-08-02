// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_twice_final_value
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 1;
System.Threading.Interlocked.Exchange(ref slot, 2);
System.Threading.Interlocked.Exchange(ref slot, 3);
__Check((slot).ToString(), "3");
