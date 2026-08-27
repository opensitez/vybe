// vybe-test: csharp/csharp_delegate_events_matrix/action_handlers_can_be_added_removed_from_events
// origin: languages/csharp/tests/csharp/test_csharp_delegate_events_matrix.rs

using static __Harness;

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

class Notifier {
    public event System.Action<int>? Raised;
    public void Raise(int value) => Raised?.Invoke(value);
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
