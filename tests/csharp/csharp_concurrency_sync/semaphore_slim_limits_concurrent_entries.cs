// vybe-test: csharp/csharp_concurrency_sync/semaphore_slim_limits_concurrent_entries
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sem=new System.Threading.SemaphoreSlim(1,1);
sem.Wait();
__Check((sem.CurrentCount).ToString(), "0");
sem.Release();
__Check((sem.CurrentCount).ToString(), "1");
