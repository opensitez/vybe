// vybe-test: csharp/csharp_lock_monitor/lock_if_branch_skips_increment
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
bool flag = false;
lock (gate) { if (flag) counter++; }
__Check((counter).ToString(), "0");
