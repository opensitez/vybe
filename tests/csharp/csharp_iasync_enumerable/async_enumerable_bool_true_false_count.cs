// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_bool_true_false_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<bool> Flags() {
    yield return true;
    yield return false;
    yield return true;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var f in Flags()) if (f) count++;
    Console.WriteLine(count);
}
Run().Wait();
