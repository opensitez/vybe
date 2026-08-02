// vybe-test: csharp/csharp_volatile_thread_memory/volatile_write_from_parameter_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int Value = 0;
    public void Set(int n) { Value = n; }
}
var box = new FlagBox();
box.Set(22);
__Check((box.Value).ToString(), "22");
