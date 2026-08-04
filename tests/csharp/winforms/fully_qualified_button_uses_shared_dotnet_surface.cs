// vybe-test: csharp/winforms/fully_qualified_button_uses_shared_dotnet_surface
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

var btn = new System.Windows.Forms.Button();
        btn.Text = "Shared";
        __P((btn.Text).ToString());
        __P((btn.__control_type).ToString());
__Check("Shared\nButton");
