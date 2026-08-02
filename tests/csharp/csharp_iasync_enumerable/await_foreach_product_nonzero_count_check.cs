// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_product_nonzero_count_check
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int product = 1;
    int count = 0;
    await foreach (var x in Stream()) { product *= x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
