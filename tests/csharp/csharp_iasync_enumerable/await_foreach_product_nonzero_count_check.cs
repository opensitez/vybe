// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_product_nonzero_count_check
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int product = 1;
    int count = 0;
    await foreach (var x in Stream()) { product *= x; count++; }
    __P((count).ToString());
}
Run().Wait();
__Check("3");

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
