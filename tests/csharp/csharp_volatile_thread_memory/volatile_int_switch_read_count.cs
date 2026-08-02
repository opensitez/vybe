// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_switch_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 3;
}
var box = new FlagBox();
int count = 0;
switch (box.Value) {
    case 3: count = 30; break;
    default: count = 0; break;
}
__Check((count).ToString(), "30");
