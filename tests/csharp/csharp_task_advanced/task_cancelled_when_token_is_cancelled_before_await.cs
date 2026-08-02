// vybe-test: csharp/csharp_task_advanced/task_cancelled_when_token_is_cancelled_before_await
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
string result="ok";
try{System.Threading.Tasks.Task.Delay(1000,cts.Token).Wait();}
catch(System.AggregateException){result="cancelled";}
__Check((result).ToString(), "cancelled");
