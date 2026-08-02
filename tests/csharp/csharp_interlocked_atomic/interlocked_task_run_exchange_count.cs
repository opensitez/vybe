// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_exchange_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int slot = 0;
var t1 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 1));
var t2 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 2));
System.Threading.Tasks.Task.WaitAll(t1, t2);
__Check((slot).ToString(), "2");
