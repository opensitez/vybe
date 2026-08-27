// vybe-test: csharp/csharp_delegates/event_basic
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

using static __Harness;

var btn = new Button();
btn.Click += () => __P(("clicked!").ToString());
btn.Press();
btn.Press();
__Check("clicked!\nclicked!");

class Button {
    public event Action Click;
    public void Press() {
        if (Click != null) Click();
    }
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
