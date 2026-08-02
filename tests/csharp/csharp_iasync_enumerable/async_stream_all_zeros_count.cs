// vybe-test: csharp/csharp_iasync_enumerable/async_stream_all_zeros_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Zeros() {
    yield return 0;
    yield return 0;
    yield return 0;
    yield return 0;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Zeros()) count++;
    Console.WriteLine(count);
}
Run().Wait();
