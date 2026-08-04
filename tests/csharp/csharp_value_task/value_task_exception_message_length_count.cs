// vybe-test: csharp/csharp_value_task/value_task_exception_message_length_count
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

async System.Threading.Tasks.ValueTask<int> Fail() {
    throw new System.Exception("err");
}
async System.Threading.Tasks.Task Run() {
    int len = 0;
    try { await Fail(); }
    catch (System.Exception ex) { len = ex.Message.Length; }
    __P((len).ToString());
}
Run().Wait();
__Check("3");
