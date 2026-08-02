// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_if_branch_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 2;
}
var box = new FlagBox();
int count = 0;
if (box.Value == 2) count = 1;
__Check((count).ToString(), "1");
