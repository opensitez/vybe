// vybe-test: csharp/csharp_volatile_thread_memory/volatile_two_fields_independent_counts
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int A = 1;
    public volatile int B = 2;
}
var box = new FlagBox();
__Check((box.A + box.B).ToString(), "3");
