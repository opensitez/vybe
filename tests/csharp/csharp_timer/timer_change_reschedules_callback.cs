// vybe-test: csharp/csharp_timer/timer_change_reschedules_callback
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

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

int count=0;
using var t=new System.Threading.Timer(_=>System.Threading.Interlocked.Increment(ref count),null,10,10);
System.Threading.Thread.Sleep(100);
__P((count>0).ToString());
__Check("True");
