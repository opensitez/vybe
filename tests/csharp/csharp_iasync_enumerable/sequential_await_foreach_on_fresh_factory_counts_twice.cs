// vybe-test: csharp/csharp_iasync_enumerable/sequential_await_foreach_on_fresh_factory_counts_twice
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int total = 0;
    await foreach (var x in Make()) total++;
    await foreach (var x in Make()) total++;
    __P((total).ToString());
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
