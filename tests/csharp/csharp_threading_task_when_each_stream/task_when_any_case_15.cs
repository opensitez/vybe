// vybe-test: csharp/csharp_threading_task_when_each_stream/task_when_any_case_15

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

var t1 = System.Threading.Tasks.Task.FromResult(15);
var t2 = System.Threading.Tasks.Task.FromResult(25);
var completed = System.Threading.Tasks.Task.WhenAny(t1, t2).Result;
__P((completed.Result == 15 || completed.Result == 25).ToString());
__Check("True");
