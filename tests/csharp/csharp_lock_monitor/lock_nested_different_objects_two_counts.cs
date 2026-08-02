// vybe-test: csharp/csharp_lock_monitor/lock_nested_different_objects_two_counts
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object outer = new object();
object inner = new object();
int counter = 0;
lock (outer) {
    counter++;
    lock (inner) { counter++; }
}
__Check((counter).ToString(), "2");
