// vybe-test: csharp/csharp_task_combinators/when_any_fast_beats_delayed_task
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

async System.Threading.Tasks.Task<int> Fast() { return 1; }
async System.Threading.Tasks.Task<int> Slow() {
    await System.Threading.Tasks.Task.Delay(1000);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Slow());
    __P((winner.Result).ToString());
}
Run().Wait();
__Check("1");
