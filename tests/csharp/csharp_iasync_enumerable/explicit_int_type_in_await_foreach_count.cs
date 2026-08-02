// vybe-test: csharp/csharp_iasync_enumerable/explicit_int_type_in_await_foreach_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 4;
    yield return 8;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (int x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
