// vybe-test: csharp/csharp_lock_monitor/monitor_is_entered_false_after_exit
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
System.Threading.Monitor.Enter(gate);
System.Threading.Monitor.Exit(gate);
__Check((System.Threading.Monitor.IsEntered(gate) ? 1 : 0).ToString(), "0");
