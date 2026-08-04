// vybe-test: csharp/csharp_task_combinators/when_all_loop_spawned_tasks_count
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

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var tasks = new System.Threading.Tasks.Task<int>[3];
    for (int i = 0; i < 3; i++) tasks[i] = N(i + 1);
    var results = await System.Threading.Tasks.Task.WhenAll(tasks);
    int count = 0;
    foreach (var x in results) count += x;
    __P((count).ToString());
}
Run().Wait();
__Check("6");
