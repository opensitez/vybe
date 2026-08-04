// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_increment_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 0;
var tasks = new System.Threading.Tasks.Task[5];
for (int i = 0; i < 5; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Increment(ref counter);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
__P((counter).ToString());
__Check("5");
