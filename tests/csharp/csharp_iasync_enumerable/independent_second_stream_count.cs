// vybe-test: csharp/csharp_iasync_enumerable/independent_second_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Pair() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Pair()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("2");

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
