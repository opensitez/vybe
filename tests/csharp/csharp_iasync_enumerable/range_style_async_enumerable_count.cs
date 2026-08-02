// vybe-test: csharp/csharp_iasync_enumerable/range_style_async_enumerable_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> FromTo(int start, int end) {
    for (int i = start; i <= end; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in FromTo(3, 8)) count++;
    Console.WriteLine(count);
}
Run().Wait();
