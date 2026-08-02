// vybe-test: csharp/csharp_lock_monitor/lock_three_sequential_increments_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
lock (gate) { counter++; }
lock (gate) { counter++; }
lock (gate) { counter++; }
__Check((counter).ToString(), "3");
