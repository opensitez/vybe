// vybe-test: csharp/csharp_lock_monitor/lock_task_run_adds_two_per_worker_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter += 2; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
