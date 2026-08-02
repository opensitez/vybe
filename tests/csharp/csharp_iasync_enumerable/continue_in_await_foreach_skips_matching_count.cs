// vybe-test: csharp/csharp_iasync_enumerable/continue_in_await_foreach_skips_matching_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 6; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        if (x % 2 == 0) continue;
        count++;
    }
    Console.WriteLine(count);
}
Run().Wait();
