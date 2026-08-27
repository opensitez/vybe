// vybe-test: csharp/csharp_design_patterns/observer_pattern_notifies_multiple_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

using static __Harness;

var btn = new Button();
btn.Click();
__Check("Clicked");

class Button {
    public event Action Clicked;
    public Button() { Clicked += () => __P("Clicked"); }
    public void Click() => Clicked?.Invoke();
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
