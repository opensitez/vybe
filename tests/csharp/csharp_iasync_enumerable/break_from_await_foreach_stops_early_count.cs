// vybe-test: csharp/csharp_iasync_enumerable/break_from_await_foreach_stops_early_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 10; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        count++;
        if (x == 2) break;
    }
    Console.WriteLine(count);
}
Run().Wait();
