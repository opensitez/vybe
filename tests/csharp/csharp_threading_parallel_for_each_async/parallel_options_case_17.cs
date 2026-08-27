// vybe-test: csharp/csharp_threading_parallel_for_each_async/parallel_options_case_17

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

var opt = new System.Threading.Tasks.ParallelOptions() { MaxDegreeOfParallelism = 2 };
__P(opt.MaxDegreeOfParallelism.ToString());
__Check("2");
