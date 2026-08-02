// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_one_equals_increment
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a = 6;
int b = 6;
System.Threading.Interlocked.Increment(ref a);
System.Threading.Interlocked.Add(ref b, 1);
__Check((a + b).ToString(), "14");
