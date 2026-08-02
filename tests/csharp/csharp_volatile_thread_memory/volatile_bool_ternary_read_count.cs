// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_ternary_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile bool Ready = true;
}
var box = new FlagBox();
int count = box.Ready ? 5 : 0;
__Check((count).ToString(), "5");
