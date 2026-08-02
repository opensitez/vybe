// vybe-test: csharp/csharp_iasync_enumerable/double_nested_async_enumerable_flat_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Inner() {
    yield return 1;
}
async System.Collections.Generic.IAsyncEnumerable<int> Middle() {
    await foreach (var x in Inner()) yield return x;
}
async System.Collections.Generic.IAsyncEnumerable<int> Outer() {
    await foreach (var x in Middle()) yield return x;
    await foreach (var x in Middle()) yield return x;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Outer()) count++;
    Console.WriteLine(count);
}
Run().Wait();
