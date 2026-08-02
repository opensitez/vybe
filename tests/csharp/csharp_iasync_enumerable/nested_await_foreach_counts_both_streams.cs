// vybe-test: csharp/csharp_iasync_enumerable/nested_await_foreach_counts_both_streams
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Inner() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> Outer() {
    await foreach (var x in Inner()) yield return x;
    await foreach (var x in Inner()) yield return x;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Outer()) count++;
    Console.WriteLine(count);
}
Run().Wait();
