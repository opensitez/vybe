// vybe-test: csharp/csharp_task_combinators/task_run_three_parallel_workers_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

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

async System.Threading.Tasks.Task Run() {
    var a = System.Threading.Tasks.Task.Run(() => 1);
    var b = System.Threading.Tasks.Task.Run(() => 2);
    var c = System.Threading.Tasks.Task.Run(() => 3);
    __P((a.Result + b.Result + c.Result).ToString());
}
Run().Wait();
__Check("6");
