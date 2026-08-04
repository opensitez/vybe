// vybe-test: csharp/csharp_value_task/value_task_bitwise_and_result
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

async System.Threading.Tasks.ValueTask<int> Mask() { return 0xF0; }
async System.Threading.Tasks.ValueTask<int> Value() { return 0xFF; }
async System.Threading.Tasks.Task Run() {
    __P((await Mask() & await Value()).ToString());
}
Run().Wait();
__Check("240");
