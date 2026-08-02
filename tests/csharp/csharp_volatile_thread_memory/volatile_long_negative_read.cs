// vybe-test: csharp/csharp_volatile_thread_memory/volatile_long_negative_read
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile long Value = -500L;
}
var box = new FlagBox();
__Check((box.Value).ToString(), "-500");
