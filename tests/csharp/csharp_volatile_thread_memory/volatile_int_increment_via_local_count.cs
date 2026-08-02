// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_increment_via_local_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 1;
}
var box = new FlagBox();
int snapshot = box.Value;
box.Value = snapshot + 2;
__Check((box.Value).ToString(), "3");
