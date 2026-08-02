// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_sums_async_stream
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int s = 0;
    await foreach (var x in Stream()) s += x;
    Console.WriteLine(s);
}
Run().Wait();
