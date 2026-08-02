// vybe-test: csharp/csharp_lock_monitor/lock_without_contention_read_then_write
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 1;
lock (gate) {
    int snapshot = counter;
    counter = snapshot + 4;
}
__Check((counter).ToString(), "5");
