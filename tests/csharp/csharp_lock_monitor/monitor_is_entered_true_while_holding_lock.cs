// vybe-test: csharp/csharp_lock_monitor/monitor_is_entered_true_while_holding_lock
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int count = 0;
System.Threading.Monitor.Enter(gate);
count = System.Threading.Monitor.IsEntered(gate) ? 1 : 0;
System.Threading.Monitor.Exit(gate);
__Check((count).ToString(), "1");
