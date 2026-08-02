// vybe-test: csharp/csharp_lock_monitor/lock_for_loop_accumulates_squares_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
for (int i = 1; i <= 4; i++) {
    lock (gate) { counter += i * i; }
}
Console.WriteLine(counter);
