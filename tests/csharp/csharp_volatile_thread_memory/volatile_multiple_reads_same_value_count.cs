// vybe-test: csharp/csharp_volatile_thread_memory/volatile_multiple_reads_same_value_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 4;
}
var box = new FlagBox();
int count = box.Value + box.Value + box.Value;
__Check((count).ToString(), "12");
