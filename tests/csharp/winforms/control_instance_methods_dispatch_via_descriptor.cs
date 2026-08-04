// vybe-test: csharp/winforms/control_instance_methods_dispatch_via_descriptor
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

var b = new Button();
        b.Show();
        b.Hide();
        __P(("methods-ok").ToString());
__Check("methods-ok");
