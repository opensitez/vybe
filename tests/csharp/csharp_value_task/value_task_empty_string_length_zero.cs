// vybe-test: csharp/csharp_value_task/value_task_empty_string_length_zero
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

async System.Threading.Tasks.ValueTask<string> Empty() { return ""; }
async System.Threading.Tasks.Task Run() {
    string s = await Empty();
    __P((s.Length).ToString());
}
Run().Wait();
__Check("0");
