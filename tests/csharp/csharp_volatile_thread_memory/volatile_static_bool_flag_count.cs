// vybe-test: csharp/csharp_volatile_thread_memory/volatile_static_bool_flag_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class FlagBox {
    public static volatile bool Done = false;
}
FlagBox.Done = true;
__Check((FlagBox.Done ? 1 : 0).ToString(), "1");
