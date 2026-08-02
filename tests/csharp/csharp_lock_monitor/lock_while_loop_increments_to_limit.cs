// vybe-test: csharp/csharp_lock_monitor/lock_while_loop_increments_to_limit
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
int n = 6;
while (n > 0) {
    lock (gate) { counter++; }
    n--;
}
Console.WriteLine(counter);
