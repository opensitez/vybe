// vybe-test: csharp/winforms/graphics_drawline_sequence_runs
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

var g = new PictureBox().CreateGraphics();
        var p = new Pen(Color.Red, 2);
        g.DrawLine(p, 0, 0, 10, 10);
        __P(("drew").ToString());
__Check("drew");
