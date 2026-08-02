// vybe-test: csharp/csharp_iasync_enumerable/interface_variable_async_enumerable_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 3;
    yield return 6;
    yield return 9;
}
async System.Threading.Tasks.Task Run() {
    System.Collections.Generic.IAsyncEnumerable<int> stream = Make();
    int count = 0;
    await foreach (var x in stream) count++;
    Console.WriteLine(count);
}
Run().Wait();
