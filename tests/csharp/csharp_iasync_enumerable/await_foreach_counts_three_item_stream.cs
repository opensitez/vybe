// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_counts_three_item_stream
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield return 3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
