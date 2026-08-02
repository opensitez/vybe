// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_add_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

int total = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Add(ref total, 2);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(total);
