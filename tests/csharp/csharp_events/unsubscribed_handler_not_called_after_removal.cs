// vybe-test: csharp/csharp_events/unsubscribed_handler_not_called_after_removal
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

using static __Harness;

int count = 0;
System.EventHandler h = (s, e) => count++;
var btn = new Button();
btn.Clicked += h;
btn.Clicked -= h;
btn.Click();
__P((count).ToString());
__Check("0");

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
