// vybe-test: csharp/csharp_threading_parallel_for_loop_state/parallel_for_case_2

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

int sum = 0;
object lockObj = new object();
System.Threading.Tasks.Parallel.For(0, 5, k => {
    lock (lockObj) { sum += k; }
});
__P(sum.ToString());
__Check("10");
