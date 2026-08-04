// vybe-test: csharp/csharp_task_advanced/task_continuation_runs_after_completion
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

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

int result=0;
System.Threading.Tasks.Task.Run(()=>7)
    .ContinueWith(t=>result=t.Result*2)
    .Wait();
__P((result).ToString());
__Check("14");
