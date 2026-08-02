// vybe-test: csharp/csharp_iasync_enumerable/immediate_yield_break_empty_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Nothing() {
    yield break;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Nothing()) count++;
    Console.WriteLine(count);
}
Run().Wait();
