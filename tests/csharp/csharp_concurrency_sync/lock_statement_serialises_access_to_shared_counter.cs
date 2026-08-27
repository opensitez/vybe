// vybe-test: csharp/csharp_concurrency_sync/lock_statement_serialises_access_to_shared_counter
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

using static __Harness;

int counter=0;
object lk=new object();
var tasks=new System.Threading.Tasks.Task[10];
for(int i=0;i<10;i++){
    tasks[i]=System.Threading.Tasks.Task.Run(()=>{lock(lk){counter++;}});
}
System.Threading.Tasks.Task.WaitAll(tasks);
__P((counter).ToString());
__Check("10");

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
