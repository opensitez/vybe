// vybe-test: csharp/csharp_concurrency_sync/monitor_try_enter_returns_false_when_already_locked
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj=new object();
System.Threading.Monitor.Enter(obj);
bool got=System.Threading.Monitor.TryEnter(obj,0);
System.Threading.Monitor.Exit(obj);
__Check((got).ToString(), "False");
