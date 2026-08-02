// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_default_zero_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value;
}
var box = new FlagBox();
__Check((box.Value).ToString(), "0");
