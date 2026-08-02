// vybe-test: csharp/csharp_volatile_thread_memory/volatile_read_after_constructor_set
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value;
    public FlagBox(int n) { Value = n; }
}
var box = new FlagBox(18);
__Check((box.Value).ToString(), "18");
