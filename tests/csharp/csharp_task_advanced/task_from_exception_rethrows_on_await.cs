// vybe-test: csharp/csharp_task_advanced/task_from_exception_rethrows_on_await
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

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

string msg="";
var t=System.Threading.Tasks.Task.FromException(new System.Exception("boom"));
try{t.Wait();}catch(System.AggregateException ae){msg=ae.InnerException.Message;}
__P((msg).ToString());
__Check("boom");
