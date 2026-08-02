// vybe-test: csharp/csharp_iasync_enumerable/nested_loop_yield_flat_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Grid() {
    for (int r = 0; r < 2; r++)
        for (int c = 0; c < 3; c++)
            yield return r * 10 + c;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Grid()) count++;
    Console.WriteLine(count);
}
Run().Wait();
