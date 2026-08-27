// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_string_length_total_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<string> Stream() {
    yield return "ab";
    yield return "cde";
    yield return "f";
}
async System.Threading.Tasks.Task Run() {
    int totalLen = 0;
    await foreach (var s in Stream()) totalLen += s.Length;
    __P((totalLen).ToString());
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
