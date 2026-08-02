// vybe-test: csharp/csharp_environment_system/gc_collect_runs_without_error
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.GC.Collect();
__Check(("ok").ToString(), "ok");
