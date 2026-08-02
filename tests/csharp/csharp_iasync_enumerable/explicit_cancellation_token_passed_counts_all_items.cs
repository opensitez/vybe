// vybe-test: csharp/csharp_iasync_enumerable/explicit_cancellation_token_passed_counts_all_items
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 1; i <= 5; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    await foreach (var x in Stream(cts.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
