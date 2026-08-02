// vybe-test: csharp/csharp_volatile_thread_memory/volatile_task_run_write_read_count
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
var box = new FlagBox();
System.Threading.Tasks.Task.Run(() => { box.Value = 6; }).Wait();
__Check((box.Value).ToString(), "6");
