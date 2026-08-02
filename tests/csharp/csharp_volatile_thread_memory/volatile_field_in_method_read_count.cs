// vybe-test: csharp/csharp_volatile_thread_memory/volatile_field_in_method_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 6;
    public int Read() { return Value; }
}
var box = new FlagBox();
__Check((box.Read()).ToString(), "6");
