// vybe-test: csharp/csharp_iasync_enumerable/task_yield_between_yields_preserves_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 10;
    await System.Threading.Tasks.Task.Yield();
    yield return 20;
    await System.Threading.Tasks.Task.Yield();
    yield return 30;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
