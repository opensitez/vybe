// vybe-test: csharp/csharp_iasync_enumerable/empty_async_enumerable_produces_zero_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Empty() {
    yield break;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Empty()) count++;
    Console.WriteLine(count);
}
Run().Wait();
