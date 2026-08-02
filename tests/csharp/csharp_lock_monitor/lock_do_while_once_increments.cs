// vybe-test: csharp/csharp_lock_monitor/lock_do_while_once_increments
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
do { lock (gate) { counter++; } } while (false);
Console.WriteLine(counter);
