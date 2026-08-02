// vybe-test: csharp/csharp_task_advanced/when_any_returns_first_completed_task
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Fast()=>await System.Threading.Tasks.Task.FromResult(1);
async System.Threading.Tasks.Task<int> Slow(){await System.Threading.Tasks.Task.Delay(1000);return 2;}
var winner=await System.Threading.Tasks.Task.WhenAny(Fast(),Slow());
__Check((winner.Result).ToString(), "1");
