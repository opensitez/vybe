// vybe-test: csharp/csharp_task_advanced/cancellation_token_is_cancelled_after_cancel_called
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
__Check((cts.Token.IsCancellationRequested).ToString(), "True");
