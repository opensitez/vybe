// vybe-test: csharp/csharp_iasync_enumerable/map_in_foreach_body_count_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    int sum = 0;
    await foreach (var x in Stream()) { sum += x * 2; count++; }
    Console.WriteLine(count);
}
Run().Wait();
