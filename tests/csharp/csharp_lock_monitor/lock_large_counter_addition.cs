// vybe-test: csharp/csharp_lock_monitor/lock_large_counter_addition
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 1000;
lock (gate) { counter += 250; }
__Check((counter).ToString(), "1250");
