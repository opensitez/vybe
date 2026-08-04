// vybe-test: csharp/csharp_async_task/exception_in_async_method_propagates_through_await
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

async System.Threading.Tasks.Task Fail() {
    await System.Threading.Tasks.Task.Yield();
    throw new System.Exception("async fail");
}
string msg = "";
try { Fail().Wait(); }
catch (System.AggregateException ae) { msg = ae.InnerException.Message; }
__P((msg).ToString());
__Check("async fail");
