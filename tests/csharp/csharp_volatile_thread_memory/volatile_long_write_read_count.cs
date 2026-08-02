// vybe-test: csharp/csharp_volatile_thread_memory/volatile_long_write_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile long Value = 0L;
}
var box = new FlagBox();
box.Value = 1000000L;
__Check((box.Value).ToString(), "1000000");
