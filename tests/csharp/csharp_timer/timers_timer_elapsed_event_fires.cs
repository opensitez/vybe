// vybe-test: csharp/csharp_timer/timers_timer_elapsed_event_fires
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

bool fired=false;
var t=new System.Timers.Timer(10){AutoReset=false};
t.Elapsed+=(_,__)=>fired=true;
t.Start();
System.Threading.Thread.Sleep(100);
__P((fired).ToString());
__Check("True");
