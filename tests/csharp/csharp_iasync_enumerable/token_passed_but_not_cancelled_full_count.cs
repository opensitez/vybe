// vybe-test: csharp/csharp_iasync_enumerable/token_passed_but_not_cancelled_full_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 9; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    await foreach (var x in Stream(cts.Token)) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("9");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
