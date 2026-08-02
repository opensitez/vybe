// vybe-test: csharp/csharp_iasync_enumerable/cancellation_token_linked_parent_not_cancelled_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 7; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var parent = new System.Threading.CancellationTokenSource();
    var child = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(parent.Token);
    int count = 0;
    await foreach (var x in Stream(child.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
