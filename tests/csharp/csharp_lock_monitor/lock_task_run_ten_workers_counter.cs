// vybe-test: csharp/csharp_lock_monitor/lock_task_run_ten_workers_counter
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[10];
for (int i = 0; i < 10; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
