// vybe-test: csharp/csharp_lock_monitor/lock_task_run_two_workers_counter
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[2];
for (int i = 0; i < 2; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
