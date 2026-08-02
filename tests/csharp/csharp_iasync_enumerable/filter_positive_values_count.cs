// vybe-test: csharp/csharp_iasync_enumerable/filter_positive_values_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Mixed() {
    yield return -1;
    yield return 2;
    yield return -3;
    yield return 4;
    yield return 5;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Mixed()) if (x > 0) count++;
    Console.WriteLine(count);
}
Run().Wait();
