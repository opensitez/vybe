// vybe-test: csharp/csharp_async_advanced/value_task_from_result_avoids_allocation
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int> GetValueAsync()=>42;
int v=await GetValueAsync();
__P((v).ToString());
__Check("42");

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
