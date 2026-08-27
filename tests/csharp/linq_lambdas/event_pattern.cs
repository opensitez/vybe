// vybe-test: csharp/linq_lambdas/event_pattern
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var btn = new Button();
btn.OnClick += () => __P(("clicked!").ToString());
btn.Click();
btn.Click();
__Check("clicked!\nclicked!");

class Button {
    public event Action OnClick;
    public void Click() { if (OnClick != null) OnClick(); }
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
