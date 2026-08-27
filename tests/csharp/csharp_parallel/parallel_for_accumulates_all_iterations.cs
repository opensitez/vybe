// vybe-test: csharp/csharp_parallel/parallel_for_accumulates_all_iterations
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

using static __Harness;

int sum=0;
System.Threading.Tasks.Parallel.For(0,100,i=>{
    System.Threading.Interlocked.Add(ref sum,i);
});
__P((sum).ToString());
__Check("4950");

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
