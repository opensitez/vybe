// vybe-test: csharp/csharp_iasync_enumerable/async_stream_negative_numbers_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Negatives() {
    yield return -1;
    yield return -2;
    yield return -3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Negatives()) count++;
    Console.WriteLine(count);
}
Run().Wait();
