// vybe-test: csharp/csharp_iasync_enumerable/cancellation_token_cancelled_before_foreach_yields_zero
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 0; i < 8; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    cts.Cancel();
    int count = 0;
    try {
        await foreach (var x in Stream(cts.Token)) count++;
    } catch (System.OperationCanceledException) { }
    Console.WriteLine(count);
}
Run().Wait();
