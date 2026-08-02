// vybe-test: csharp/csharp_volatile_thread_memory/volatile_static_int_field_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public static volatile int Shared = 0;
}
FlagBox.Shared = 12;
__Check((FlagBox.Shared).ToString(), "12");
