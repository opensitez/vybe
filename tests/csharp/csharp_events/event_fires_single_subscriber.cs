// vybe-test: csharp/csharp_events/event_fires_single_subscriber
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

class Button {
    public event System.EventHandler Clicked;
    public void Click() => Clicked?.Invoke(this, System.EventArgs.Empty);
}
int count = 0;
var btn = new Button();
btn.Clicked += (s, e) => count++;
btn.Click();
__P((count).ToString());
__Check("1");
