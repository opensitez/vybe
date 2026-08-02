// vybe-test: csharp/csharp_iasync_enumerable/sequential_await_foreach_on_fresh_factory_counts_twice
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int total = 0;
    await foreach (var x in Make()) total++;
    await foreach (var x in Make()) total++;
    Console.WriteLine(total);
}
Run().Wait();
