// vybe-test: csharp/csharp_delegates/event_basic
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

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
    public event Action Click;
    public void Press() {
        if (Click != null) Click();
    }
}
var btn = new Button();
btn.Click += () => __P(("clicked!").ToString());
btn.Press();
btn.Press();
__Check("clicked!\nclicked!");
