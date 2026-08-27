// vybe-test: csharp/csharp_parallel/parallel_for_each_processes_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

using static __Harness;

var items=new[]{1,2,3,4,5}
;
int sum=0;
System.Threading.Tasks.Parallel.ForEach(items,n=>{
    System.Threading.Interlocked.Add(ref sum,n);
});
__P((sum).ToString());
__Check("15");

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
