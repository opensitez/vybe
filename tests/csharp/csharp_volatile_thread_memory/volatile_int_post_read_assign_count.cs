// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_post_read_assign_count
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
int first = box.Value;
box.Value = 2;
int second = box.Value;
__Check((first + second).ToString(), "3");
