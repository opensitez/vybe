// vybe-test: csharp/csharp_lock_monitor/monitor_try_enter_fails_when_already_locked
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
System.Threading.Monitor.Enter(gate);
bool got = System.Threading.Monitor.TryEnter(gate, 0);
System.Threading.Monitor.Exit(gate);
__Check((got ? 1 : 0).ToString(), "0");
