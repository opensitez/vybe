// vybe-test: csharp/csharp_iasync_enumerable/take_until_threshold_reached_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 1; i <= 10; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        count++;
        if (x >= 5) break;
    }
    Console.WriteLine(count);
}
Run().Wait();
