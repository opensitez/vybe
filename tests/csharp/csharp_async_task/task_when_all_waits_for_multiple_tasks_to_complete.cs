// vybe-test: csharp/csharp_async_task/task_when_all_waits_for_multiple_tasks_to_complete
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

async System.Threading.Tasks.Task<int> Val(int n) {
    await System.Threading.Tasks.Task.Yield();
    return n;
}
var results = System.Threading.Tasks.Task.WhenAll(Val(1), Val(2), Val(3)).Result;
int sum = 0;
foreach (var r in results) sum += r;
Console.WriteLine(sum);
