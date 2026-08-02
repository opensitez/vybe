// vybe-test: csharp/csharp_iasync_enumerable/odd_only_conditional_yield_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Odds(int n) {
    for (int i = 0; i < n; i++)
        if (i % 2 == 1) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Odds(8)) count++;
    Console.WriteLine(count);
}
Run().Wait();
