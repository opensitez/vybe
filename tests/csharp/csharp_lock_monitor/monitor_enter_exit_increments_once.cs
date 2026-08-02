// vybe-test: csharp/csharp_lock_monitor/monitor_enter_exit_increments_once
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
System.Threading.Monitor.Enter(gate);
counter++;
System.Threading.Monitor.Exit(gate);
__Check((counter).ToString(), "1");
