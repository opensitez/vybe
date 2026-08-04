// vybe-test: csharp/csharp_task_combinators/when_all_task_run_four_workers_sum
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
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.Run(() => 1),
        System.Threading.Tasks.Task.Run(() => 2),
        System.Threading.Tasks.Task.Run(() => 3),
        System.Threading.Tasks.Task.Run(() => 4)
    );
    int count = 0;
    foreach (var x in results) count += x;
    __P((count).ToString());
}
Run().Wait();
__Check("10");
