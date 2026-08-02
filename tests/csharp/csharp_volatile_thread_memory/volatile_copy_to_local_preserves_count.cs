// vybe-test: csharp/csharp_volatile_thread_memory/volatile_copy_to_local_preserves_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 8;
}
var box = new FlagBox();
int local = box.Value;
__Check((local).ToString(), "8");
