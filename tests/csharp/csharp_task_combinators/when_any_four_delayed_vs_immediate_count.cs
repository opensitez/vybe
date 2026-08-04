// vybe-test: csharp/csharp_task_combinators/when_any_four_delayed_vs_immediate_count
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

async System.Threading.Tasks.Task<int> Now() { return 2; }
async System.Threading.Tasks.Task<int> Later() {
    await System.Threading.Tasks.Task.Delay(200);
    return 3;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Now(), Later(), Later(), Later());
    __P((winner.Result).ToString());
}
Run().Wait();
__Check("2");
