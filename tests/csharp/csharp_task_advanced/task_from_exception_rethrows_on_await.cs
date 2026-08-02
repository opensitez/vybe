// vybe-test: csharp/csharp_task_advanced/task_from_exception_rethrows_on_await
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string msg="";
var t=System.Threading.Tasks.Task.FromException(new System.Exception("boom"));
try{t.Wait();}catch(System.AggregateException ae){msg=ae.InnerException.Message;}
__Check((msg).ToString(), "boom");
