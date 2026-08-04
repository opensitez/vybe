// vybe-test: csharp/csharp_value_task/value_task_chain_two_methods
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

async System.Threading.Tasks.ValueTask<int> A() { return 2; }
async System.Threading.Tasks.ValueTask<int> B(int x) { return x + 3; }
async System.Threading.Tasks.Task Run() {
    __P((await B(await A())).ToString());
}
Run().Wait();
__Check("5");
