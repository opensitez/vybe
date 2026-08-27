// vybe-test: csharp/csharp_iasync_enumerable/nested_await_foreach_counts_both_streams
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Inner() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> Outer() {
    await foreach (var x in Inner()) yield return x;
    await foreach (var x in Inner()) yield return x;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Outer()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("4");

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
