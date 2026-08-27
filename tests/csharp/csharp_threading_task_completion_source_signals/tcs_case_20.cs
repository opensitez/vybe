// vybe-test: csharp/csharp_threading_task_completion_source_signals/tcs_case_20

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var tcs = new System.Threading.Tasks.TaskCompletionSource<int>();
tcs.SetResult(20);
__P(tcs.Task.Result.ToString());
__Check("20");
