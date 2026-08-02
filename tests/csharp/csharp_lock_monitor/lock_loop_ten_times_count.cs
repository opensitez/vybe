// vybe-test: csharp/csharp_lock_monitor/lock_loop_ten_times_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
for (int i = 0; i < 10; i++) lock (gate) { counter++; }
Console.WriteLine(counter);
