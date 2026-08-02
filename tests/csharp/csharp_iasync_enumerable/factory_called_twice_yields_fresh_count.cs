// vybe-test: csharp/csharp_iasync_enumerable/factory_called_twice_yields_fresh_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Fresh() {
    yield return 100;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Fresh()) count++;
    await foreach (var x in Fresh()) count++;
    Console.WriteLine(count);
}
Run().Wait();
