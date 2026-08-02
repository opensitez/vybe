// vybe-test: csharp/csharp_iasync_enumerable/conditional_yield_only_evens_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Evens(int max) {
    for (int i = 0; i < max; i++)
        if (i % 2 == 0) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Evens(7)) count++;
    Console.WriteLine(count);
}
Run().Wait();
