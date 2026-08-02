// vybe-test: csharp/winforms/graphics_drawline_sequence_runs
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var g = new PictureBox().CreateGraphics();
        var p = new Pen(Color.Red, 2);
        g.DrawLine(p, 0, 0, 10, 10);
        __Check(("drew").ToString(), "drew");
