// vybe-test: csharp/csharp_parallel/parallel_invoke_runs_all_actions
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

using static __Harness;

int a=0,b=0;
System.Threading.Tasks.Parallel.Invoke(
    ()=>System.Threading.Interlocked.Exchange(ref a,1),
    ()=>System.Threading.Interlocked.Exchange(ref b,2)
);
__P((a+b).ToString());
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
