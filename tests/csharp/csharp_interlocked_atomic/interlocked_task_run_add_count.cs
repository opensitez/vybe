// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_add_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int total = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Add(ref total, 2);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
__P((total).ToString());
__Check("8");

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
