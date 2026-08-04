// vybe-test: csharp/csharp_iasync_enumerable/cancellation_token_linked_parent_not_cancelled_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

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
    __P((count).ToString());
}
Run().Wait();
__Check("7");
