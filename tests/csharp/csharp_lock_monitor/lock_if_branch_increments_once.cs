// vybe-test: csharp/csharp_lock_monitor/lock_if_branch_increments_once
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
bool flag = true;
lock (gate) { if (flag) counter++; }
__Check((counter).ToString(), "1");
