// vybe-test: csharp/csharp_async_task/async_method_returns_task_t_with_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

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

async System.Threading.Tasks.Task<int> Compute() {
    await System.Threading.Tasks.Task.Yield();
    return 7;
}
__P((Compute().Result).ToString());
__Check("7");
