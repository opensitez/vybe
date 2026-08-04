// vybe-test: csharp/csharp_timer/threading_timer_fires_callback_after_delay
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
using var t=new System.Threading.Timer(_=>{fired=true;},null,10,System.Threading.Timeout.Infinite);
System.Threading.Thread.Sleep(100);
__P((fired).ToString());
__Check("True");
