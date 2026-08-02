// vybe-test: csharp/csharp_volatile_thread_memory/volatile_read_in_expression_sum
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile int X = 2;
    public volatile int Y = 3;
}
var box = new FlagBox();
__Check((box.X + box.Y).ToString(), "5");
