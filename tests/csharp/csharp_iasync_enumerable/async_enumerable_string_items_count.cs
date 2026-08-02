// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_string_items_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<string> Words() {
    yield return "a";
    yield return "bb";
    yield return "ccc";
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var w in Words()) count++;
    Console.WriteLine(count);
}
Run().Wait();
