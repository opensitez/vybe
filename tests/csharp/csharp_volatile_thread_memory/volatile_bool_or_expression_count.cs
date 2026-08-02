// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_or_expression_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public volatile bool A = false;
    public volatile bool B = true;
}
var box = new FlagBox();
__Check(((box.A || box.B) ? 1 : 0).ToString(), "1");
