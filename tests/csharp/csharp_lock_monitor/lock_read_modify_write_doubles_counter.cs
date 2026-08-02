// vybe-test: csharp/csharp_lock_monitor/lock_read_modify_write_doubles_counter
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 2;
lock (gate) { counter = counter * 2; }
__Check((counter).ToString(), "4");
