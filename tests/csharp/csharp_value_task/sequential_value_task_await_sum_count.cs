// vybe-test: csharp/csharp_value_task/sequential_value_task_await_sum_count
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

async System.Threading.Tasks.ValueTask<int> N(int x) { return x; }
async System.Threading.Tasks.Task Run() {
    int total = await N(1) + await N(2) + await N(3);
    __P((total).ToString());
}
Run().Wait();
__Check("6");
