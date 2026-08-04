// vybe-test: csharp/csharp_events/multiple_subscribers_all_called_in_order
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

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

class Emitter {
    public event System.Action<string> Signal;
    public void Emit(string v) => Signal?.Invoke(v);
}
string log = "";
var e = new Emitter();
e.Signal += v => log += "A";
e.Signal += v => log += "B";
e.Emit("x");
__P((log).ToString());
__Check("AB");
