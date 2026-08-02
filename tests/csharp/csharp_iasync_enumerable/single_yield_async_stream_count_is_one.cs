// vybe-test: csharp/csharp_iasync_enumerable/single_yield_async_stream_count_is_one
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> One() {
    yield return 99;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in One()) count++;
    Console.WriteLine(count);
}
Run().Wait();
