// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_long_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<long> Longs() {
    yield return 100L;
    yield return 200L;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var v in Longs()) count++;
    Console.WriteLine(count);
}
Run().Wait();
