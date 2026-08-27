// vybe-test: csharp/csharp_threading_task_when_each_stream/task_when_any_case_8

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

var t1 = System.Threading.Tasks.Task.FromResult(8);
var t2 = System.Threading.Tasks.Task.FromResult(18);
var completed = System.Threading.Tasks.Task.WhenAny(t1, t2).Result;
__P((completed.Result == 8 || completed.Result == 18).ToString());
__Check("True");
