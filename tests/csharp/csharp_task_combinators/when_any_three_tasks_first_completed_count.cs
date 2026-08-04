// vybe-test: csharp/csharp_task_combinators/when_any_three_tasks_first_completed_count
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

async System.Threading.Tasks.Task<int> A() { return 10; }
async System.Threading.Tasks.Task<int> B() { return 20; }
async System.Threading.Tasks.Task<int> C() { return 30; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B(), C());
    __P((winner.Result).ToString());
}
Run().Wait();
__Check("10");
