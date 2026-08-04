// vybe-test: csharp/csharp_value_task/value_task_subtraction_chain_count
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

async System.Threading.Tasks.ValueTask<int> Start() { return 50; }
async System.Threading.Tasks.ValueTask<int> Take() { return 8; }
async System.Threading.Tasks.Task Run() {
    __P((await Start() - await Take()).ToString());
}
Run().Wait();
__Check("42");
