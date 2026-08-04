// vybe-test: csharp/csharp_events/custom_event_args_carries_data_to_handler
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

class DataArgs : System.EventArgs { public int Value; }
class Source {
    public event System.EventHandler<DataArgs> Changed;
    public void Change(int v) => Changed?.Invoke(this, new DataArgs{Value=v});
}
int received = 0;
var src = new Source();
src.Changed += (s, e) => received = e.Value;
src.Change(77);
__P((received).ToString());
__Check("77");
