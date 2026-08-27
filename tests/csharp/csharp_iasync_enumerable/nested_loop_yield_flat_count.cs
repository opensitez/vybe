// vybe-test: csharp/csharp_iasync_enumerable/nested_loop_yield_flat_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Grid() {
    for (int r = 0; r < 2; r++)
        for (int c = 0; c < 3; c++)
            yield return r * 10 + c;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Grid()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("6");

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
