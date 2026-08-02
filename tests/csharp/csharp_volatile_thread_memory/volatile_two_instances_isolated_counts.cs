// vybe-test: csharp/csharp_volatile_thread_memory/volatile_two_instances_isolated_counts
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
var a = new FlagBox();
var b = new FlagBox();
a.Value = 4;
b.Value = 5;
__Check((a.Value + b.Value).ToString(), "9");
