// vybe-test: csharp/csharp_lock_monitor/lock_zero_initial_increment_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
lock (gate) { counter = counter + 0 + 1; }
__Check((counter).ToString(), "1");
