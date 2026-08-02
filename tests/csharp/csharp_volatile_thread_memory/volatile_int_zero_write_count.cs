// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_zero_write_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 99;
}
var box = new FlagBox();
box.Value = 0;
__Check((box.Value).ToString(), "0");
