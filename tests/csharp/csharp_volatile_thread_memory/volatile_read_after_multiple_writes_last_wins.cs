// vybe-test: csharp/csharp_volatile_thread_memory/volatile_read_after_multiple_writes_last_wins
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 1;
box.Value = 2;
box.Value = 9;
__Check((box.Value).ToString(), "9");
