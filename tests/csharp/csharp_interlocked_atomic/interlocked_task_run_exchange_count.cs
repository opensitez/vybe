// vybe-test: csharp/csharp_interlocked_atomic/interlocked_task_run_exchange_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

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

int slot = 0;
var t1 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 1));
var t2 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 2));
System.Threading.Tasks.Task.WaitAll(t1, t2);
__P((slot).ToString());
__Check("2");
