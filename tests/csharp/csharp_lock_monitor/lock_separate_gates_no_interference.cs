// vybe-test: csharp/csharp_lock_monitor/lock_separate_gates_no_interference
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object g1 = new object();
object g2 = new object();
int c1 = 0;
int c2 = 0;
lock (g1) { c1 = 3; }
lock (g2) { c2 = 4; }
__Check((c1 + c2).ToString(), "7");
