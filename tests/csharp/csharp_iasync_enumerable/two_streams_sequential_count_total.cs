// vybe-test: csharp/csharp_iasync_enumerable/two_streams_sequential_count_total
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> A() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> B() {
    yield return 10;
    yield return 20;
    yield return 30;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in A()) count++;
    await foreach (var x in B()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("5");

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
