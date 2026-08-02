// vybe-test: csharp/csharp_environment_system/gc_get_total_memory_returns_positive_long
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.GC.GetTotalMemory(false)>0).ToString(), "True");
