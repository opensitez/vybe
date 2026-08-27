// vybe-test: csharp/csharp_iasync_enumerable/skip_first_two_via_flag_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 6; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int seen = 0;
    int count = 0;
    await foreach (var x in Stream()) {
        seen++;
        if (seen <= 2) continue;
        count++;
    }
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
