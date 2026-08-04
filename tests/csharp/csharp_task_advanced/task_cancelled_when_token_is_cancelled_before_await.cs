// vybe-test: csharp/csharp_task_advanced/task_cancelled_when_token_is_cancelled_before_await
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

var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
string result="ok";
try{System.Threading.Tasks.Task.Delay(1000,cts.Token).Wait();}
catch(System.AggregateException){result="cancelled";}
__P((result).ToString());
__Check("cancelled");
