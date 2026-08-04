// vybe-test: csharp/csharp_task_combinators/continue_with_doubles_task_result
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
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => 7)
        .ContinueWith(t => count = t.Result * 2);
    __P((count).ToString());
}
Run().Wait();
__Check("14");
