// vybe-test: csharp/winforms/control_create_graphics_returns_graphics_handle
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

var pb = new PictureBox();
        pb.Name = "art";
        var g = pb.CreateGraphics();
        __P((g == null ? "null-graphics" : "have-graphics").ToString());
__Check("have-graphics");
