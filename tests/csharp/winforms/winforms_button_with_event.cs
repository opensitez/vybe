// vybe-test: csharp/winforms/winforms_button_with_event
// origin: languages/csharp/tests/csharp/test_winforms.rs

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

var btn = new Button();
        btn.Name = "btn1";
        btn.Text = "Click";
        __P((btn.Name).ToString());
        __P((btn.Text).ToString());
        __P((btn.__control_type).ToString());
__Check("btn1\nClick\nButton");
