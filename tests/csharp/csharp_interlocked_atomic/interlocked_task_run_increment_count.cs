// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_increment_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int counter = 0;
var tasks = new System.Threading.Tasks.Task[5];
for (int i = 0; i < 5; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Increment(ref counter);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
