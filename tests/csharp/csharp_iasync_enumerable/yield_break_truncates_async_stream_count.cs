// vybe-test: csharp/csharp_iasync_enumerable/yield_break_truncates_async_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield break;
    yield return 99;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
