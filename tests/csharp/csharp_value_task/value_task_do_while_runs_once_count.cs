// vybe-test: csharp/csharp_value_task/value_task_do_while_runs_once_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

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

async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    do { count += await Step(); } while (false);
    __P((count).ToString());
}
Run().Wait();
__Check("1");
