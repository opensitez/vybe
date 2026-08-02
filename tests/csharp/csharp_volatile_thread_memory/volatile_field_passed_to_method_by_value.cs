// vybe-test: csharp/csharp_volatile_thread_memory/volatile_field_passed_to_method_by_value
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Double(int n) { return n * 2; }
class FlagBox {
    public volatile int Value = 5;
}
var box = new FlagBox();
__Check((Double(box.Value)).ToString(), "10");
