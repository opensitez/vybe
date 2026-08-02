// vybe-test: csharp/csharp_lock_monitor/lock_two_objects_independent_counters
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object a = new object();
object b = new object();
int ca = 0;
int cb = 0;
lock (a) { ca++; }
lock (b) { cb += 2; }
__Check((ca + cb).ToString(), "3");
