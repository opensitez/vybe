// vybe-test: csharp/csharp_async_advanced/async_stream_yields_values_to_await_foreach
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

using static __Harness;

async System.Collections.Generic.IAsyncEnumerable<int> Sequence(){
    for(int i=1;i<=3;i++){
        await System.Threading.Tasks.Task.Yield();
        yield return i;
    }
}
int sum=0;
await foreach(var n in Sequence()) sum+=n;
__P((sum).ToString());
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
