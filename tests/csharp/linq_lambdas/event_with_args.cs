// vybe-test: csharp/linq_lambdas/event_with_args
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

class Timer {
    public event Action<int> OnTick;
    public void Tick(int count) { if (OnTick != null) OnTick(count); }
}
var t = new Timer();
t.OnTick += n => __P(("tick " + n).ToString());
t.Tick(1);
t.Tick(2);
__Check("tick 1\ntick 2");
