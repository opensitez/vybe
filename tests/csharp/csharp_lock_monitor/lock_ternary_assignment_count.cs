// vybe-test: csharp/csharp_lock_monitor/lock_ternary_assignment_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
bool pick = true;
lock (gate) { counter = pick ? 3 : 8; }
__Check((counter).ToString(), "3");
