// vybe-test: csharp/csharp_lock_monitor/lock_task_run_two_gates_isolated_totals
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
var t1 = System.Threading.Tasks.Task.Run(() => { lock (g1) { c1 += 5; } });
var t2 = System.Threading.Tasks.Task.Run(() => { lock (g2) { c2 += 6; } });
System.Threading.Tasks.Task.WaitAll(t1, t2);
__Check((c1 + c2).ToString(), "11");
