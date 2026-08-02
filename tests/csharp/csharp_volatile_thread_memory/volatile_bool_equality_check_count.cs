// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_equality_check_count
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
__Check((box.Ready == true ? 1 : 0).ToString(), "1");
