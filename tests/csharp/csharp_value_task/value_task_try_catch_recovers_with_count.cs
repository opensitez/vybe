// vybe-test: csharp/csharp_value_task/value_task_try_catch_recovers_with_count
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

async System.Threading.Tasks.ValueTask<int> Risky(bool fail) {
    if (fail) throw new System.Exception("no");
    return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    try { count = await Risky(true); }
    catch (System.Exception) { count = 2; }
    __P((count).ToString());
}
Run().Wait();
__Check("2");
