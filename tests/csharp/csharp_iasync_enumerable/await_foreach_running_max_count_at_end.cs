// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_running_max_count_at_end
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 3;
    yield return 7;
    yield return 5;
    yield return 9;
}
async System.Threading.Tasks.Task Run() {
    int max = int.MinValue;
    int count = 0;
    await foreach (var x in Stream()) { if (x > max) max = x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
