// vybe-test: csharp/csharp_delegate_events_matrix/action_handlers_can_be_added_removed_from_events
// origin: languages/csharp/tests/csharp/test_csharp_delegate_events_matrix.rs

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

class Notifier {
    public event System.Action<int>? Raised;
    public void Raise(int value) => Raised?.Invoke(value);
}

int total = 0;
var notifier = new Notifier();
System.Action<int> add = value => total += value;
System.Action<int> sub = value => total -= value;
notifier.Raised += add;
notifier.Raised += sub;
notifier.Raise(5);
notifier.Raised += add;
notifier.Raise(3);
notifier.Raised -= sub;
notifier.Raise(2);
__P((total).ToString());
__Check("7");
