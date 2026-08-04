// vybe-test: csharp/csharp_value_task/value_task_multiply_two_results
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

async System.Threading.Tasks.ValueTask<int> Left() { return 6; }
async System.Threading.Tasks.ValueTask<int> Right() { return 7; }
async System.Threading.Tasks.Task Run() {
    __P((await Left() * await Right()).ToString());
}
Run().Wait();
__Check("42");
