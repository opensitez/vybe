// vybe-test: csharp/csharp_task_combinators/task_run_bool_false_count
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
    var t = System.Threading.Tasks.Task.Run(() => false);
    __P((t.Result ? 1 : 0).ToString());
}
Run().Wait();
__Check("0");
