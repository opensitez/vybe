// vybe-test: csharp/csharp_iasync_enumerable/async_stream_loop_twenty_items_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Range20() {
    for (int i = 0; i < 20; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Range20()) count++;
    Console.WriteLine(count);
}
Run().Wait();
