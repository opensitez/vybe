// vybe-test: csharp/csharp_lock_monitor/lock_preserves_counter_when_no_contention
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 7;
lock (gate) { counter += 3; }
__Check((counter).ToString(), "10");
