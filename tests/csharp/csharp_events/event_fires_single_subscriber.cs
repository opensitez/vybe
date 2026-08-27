// vybe-test: csharp/csharp_events/event_fires_single_subscriber
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

using static __Harness;

int count = 0;
var btn = new Button();
btn.Clicked += (s, e) => count++;
btn.Click();
__P((count).ToString());
__Check("1");

class Button {
    public event System.EventHandler Clicked;
    public void Click() => Clicked?.Invoke(this, System.EventArgs.Empty);
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
