// vybe-test: csharp/winforms/form_subclass_constructs_via_descriptor_and_inherits_properties
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

class MyForm : Form {
        }
        var f = new MyForm();
        f.Text = "hello";
        __P((f.Text).ToString());
        __P((f.__control_type).ToString());
__Check("hello\nForm");
