// vybe-test: csharp/csharp_iasync_enumerable/async_stream_with_await_before_first_yield_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> DelayedStart() {
    await System.Threading.Tasks.Task.Yield();
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in DelayedStart()) count++;
    Console.WriteLine(count);
}
Run().Wait();
